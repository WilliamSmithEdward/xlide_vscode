// Rule family: parentheses the VBE reads otherwise, or refuses (issue #236).
// Measured in Excel 16.0 (build 20326, 2026-09-30) with a full compile.
//
// A paren that follows a name, `)` or `]` calls or indexes; any other paren
// groups an expression, and so does one after a statement's callee with a
// space between, `Take (c)`.
//
//  - collection-operand: an object in grouping parentheses is its default
//    member, and a Collection's Item needs an index: `Set d = (c)`,
//    `Set d = (New Collection)`, `c.Add (c)`, `Take (c)`, `Call Take((c))`,
//    `With (New Collection)` -> "Argument not optional".
//  - malformed-statement: `(Range("A1")).Address` and `(Application).Name`,
//    a member of a grouping paren -> "Syntax error" ("Invalid or unqualified
//    reference" after Print); `v = ()`; `UBound((a))`; `G((b:=1), a:=2)`, a
//    named argument inside one; `Mid((s), 1, 1) = "x"`; `(n) = 1`, a
//    statement that starts with one -> "Syntax error".

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { MemberCompletionContext } from '../../completion/memberAccess';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { ModuleNode, Span } from '../../parser/nodes';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { PushFn } from '../analysisContext';
import { normalizeType, createObjectDefaultQueries, typeEnvironmentFor } from '../typeInference';
import {
	activeModuleMembers,
	firstExecutableTokenIndex,
	forEachStatementWithHeaders,
	matchParenFrom,
	statementTokens,
	tokenName,
	tokenText,
} from '../walker';

/** Keywords that an expression can follow: a paren after one groups. */
const OPERAND_KEYWORDS: ReadonlySet<string> = new Set([
	'and', 'or', 'xor', 'eqv', 'imp', 'not', 'mod', 'like', 'is', 'to', 'step', 'then', 'else', 'elseif',
	'if', 'while', 'until', 'case', 'print', 'call', 'return', 'with', 'in', 'each', 'set', 'let',
	'select', 'do', 'loop', 'for', 'redim', 'erase', 'typeof', 'new',
]);

const MID_STATEMENTS: ReadonlySet<string> = new Set(['mid', 'mid$', 'midb', 'midb$']);

export function checkParentheses(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	memberCtx: MemberCompletionContext,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const defaultQueries = createObjectDefaultQueries(memberCtx);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const env = typeEnvironmentFor(symbols, member);
		const needsIndex = (toks: readonly VbaToken[], from: number, to: number): string | undefined => {
			// `New Collection`, or a variable of such a type.
			const type = to - from === 2 && tokenText(toks[from]) === 'new' ? tokenName(toks[from + 1])
				: to - from === 1 ? env.get(tokenName(toks[from])?.toLowerCase() ?? '') : undefined;
			return type && defaultQueries.needsIndex(type) ? type : undefined;
		};
		forEachStatementWithHeaders(source, member.body, (stmt) => {
			checkStatement(source, stmt.span, needsIndex, push);
		}, activity);
	}
}

function checkStatement(
	source: string,
	span: Span,
	needsIndex: (toks: readonly VbaToken[], from: number, to: number) => string | undefined,
	push: PushFn,
): void {
	const all = statementTokens(source, span);
	const first = firstExecutableTokenIndex(all);
	const toks = all.slice(first);
	if (toks.length === 0) {
		return;
	}
	const at = (tok: VbaToken): Span => ({ start: span.start + tok.start, end: span.start + tok.end });
	const syntax = (message: string, where: Span, error = 'Syntax error'): void => {
		push('malformedStatement', `${message}. This is a VBE compile error: ${error}.`, where);
	};
	if (toks[0].rawText === '(') {
		syntax('A statement cannot start with a parenthesis', at(toks[0]));
		return;
	}
	const head = tokenText(toks[0]);
	// Where a statement's callee ends: `Take (c)` and `c.Add (c)` group their
	// argument; `Take(c)` in an expression calls.
	let callee = 0;
	while (toks[callee + 1]?.rawText === '.' && tokenName(toks[callee + 2])) {
		callee += 2;
	}
	const grouping = (i: number): boolean => {
		let prev = toks[i - 1];
		if (!prev) {
			return true;
		}
		// `Mid$(`: a type character glued to a name is part of the name.
		if (/^[$%&!#@]$/.test(prev.rawText) && toks[i - 2] && toks[i - 2].end === prev.start && tokenName(toks[i - 2])) {
			prev = toks[i - 2];
		}
		if (i - 1 === callee && prev.end < toks[i].start && tokenName(prev)) {
			return true;
		}
		if (prev.rawText === ')' || prev.rawText === ']' || prev.kind === 'identifier' || prev.kind === 'bracketedIdentifier') {
			return false;
		}
		return prev.kind !== 'keyword' || OPERAND_KEYWORDS.has(tokenText(prev));
	};
	// A single group is already linear; avoid index allocation on the common path.
	const firstOpen = toks.findIndex(tok => tok.rawText === '(');
	if (firstOpen < 0) { return; }
	const multiple = toks.some((tok, index) => index > firstOpen && tok.rawText === '(');
	const facts = multiple ? parenthesisFacts(toks) : undefined;
	for (let i = firstOpen; i < toks.length; i++) {
		if (toks[i].rawText !== '(') {
			continue;
		}
		const close = facts ? facts.closes[i] : matchParenFrom(toks, i);
		if (close < 0) {
			return;
		}
		if (!grouping(i)) {
			// `UBound((a))`, `UBound((a) + 0)`: the array must be a name.
			const name = tokenText(toks[i - 1]);
			if ((name === 'ubound' || name === 'lbound') && toks[i + 1]?.rawText === '(') {
				syntax(`${toks[i - 1].rawText} takes an array by its name, not in parentheses`, at(toks[i + 1]));
			}
			// `Mid((s), 1, 1) = "x"`: the Mid statement writes into a variable.
			if (i === 1 && MID_STATEMENTS.has(head) && toks[i + 1]?.rawText === '(') {
				syntax('The Mid statement writes into a variable, not a value in parentheses', at(toks[i + 1]));
			}
			continue;
		}
		if (close === i + 1) {
			syntax('Empty parentheses hold no value', at(toks[i]));
			continue;
		}
		const named = facts ? facts.nextNamed[i + 1] : firstNamedArgument(toks, i + 1, close);
		// Without facts there is only one opening paren, so its contents
		// cannot contain a deeper group before this matching close.
		if (named < close && (!facts || facts.depths[named] === facts.depths[i] + 1)) {
			syntax('A named argument cannot stand inside parentheses', at(toks[named]));
		}
		if (toks[close + 1]?.rawText === '.') {
			const afterPrint = tokenText(toks[i - 1]) === 'print';
			syntax(
				afterPrint
					? 'A member cannot be read from a value in parentheses: after Print this reads as a With member'
					: 'A member cannot be read from a value in parentheses',
				at(toks[close + 1]),
				afterPrint ? 'Invalid or unqualified reference' : 'Syntax error',
			);
			continue;
		}
		// A grouping paren always evaluates what it holds: an object whose
		// default member needs an index has no value there.
		const type = needsIndex(toks, i + 1, close);
		if (type) {
			push(
				'collectionOperand',
				`'${source.slice(span.start + toks[i].start, span.start + toks[close].end)}' is the default member of ${article(type)} ${type}, whose Item needs an index. This is a VBE compile error: Argument not optional.`,
				{ start: span.start + toks[i].start, end: span.start + toks[close].end },
			);
		}
	}
}

/** Match groups and locate their first named token without rescanning nested slices. */
function parenthesisFacts(toks: readonly VbaToken[]) {
	const closes = new Int32Array(toks.length).fill(-1);
	const depths = new Int32Array(toks.length);
	const nextNamed = new Int32Array(toks.length + 1).fill(toks.length);
	const stack: number[] = [];
	let depth = 0;
	for (let i = 0; i < toks.length; i++) {
		depths[i] = depth;
		if (toks[i].rawText === '(') { stack.push(i); depth++; }
		else if (toks[i].rawText === ')') {
			depth--;
			const open = stack.pop();
			if (open !== undefined) { closes[open] = i; }
		}
	}
	for (let i = toks.length - 1; i >= 0; i--) {
		nextNamed[i] = toks[i].rawText === ':=' ? i : nextNamed[i + 1];
	}
	return { closes, depths, nextNamed };
}

function firstNamedArgument(toks: readonly VbaToken[], from: number, to: number): number {
	for (let i = from; i < to; i++) { if (toks[i].rawText === ':=') { return i; } }
	return to;
}
function article(type: string): string {
	return /^[aeiou]/i.test(normalizeType(type) ?? type) ? 'an' : 'a';
}
