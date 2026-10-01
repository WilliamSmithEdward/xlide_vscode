// Rule family: statement forms the VBE refuses while compiling (issue #125).
// Measured in Excel 16.0 (build 20326, 2026-09-25):
//
//  - collection-operand: `x = c + 1` with c As New Collection -> "Argument
//    not optional". A Collection's default member Item takes an index, so
//    the bare variable has no value for the operator.
//  - sub-used-as-value: `x = Foo` where Foo is a Sub -> "Expected Function
//    or variable".
//  - rem-after-then: `If x Then Rem note` -> "Syntax error". Rem starts a
//    comment only at the start of a statement.
//  - rem-after-statement (issue #231): `x = 1 Rem note`, `Next Rem note`
//    -> "Syntax error"; `Dim m As Long Rem note` at module level -> "Expected:
//    end of statement". In a one-line If's Then or Else list it is a comment.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { ModuleNode, Span } from '../../parser/nodes';
import { isDecimalLineNumber } from '../../lexer/tokenHelpers';
import type { VbaToken } from '../../lexer/tokenKinds';
import { tokenizeCached } from '../../lexer/tokenize';
import { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { VbaProcedureSignature } from '../../symbols/symbolModel';
import { procedureSymbolFor, type PushFn } from '../analysisContext';
import { isKnownScalarType, normalizeType, objectValueNeedsIndex, typeEnvironmentFor } from '../typeInference';
import type { MemberCompletionContext } from '../../completion/memberAccess';
import {
	activeModuleMembers,
	bareAssignmentTarget,
	firstExecutableTokenIndex,
	forEachStatement,
	statementAndBranchSpans,
	statementTokens,
	tokenName,
	tokenText,
} from '../walker';

const SCALAR_OPERATORS: ReadonlySet<string> = new Set(['=', '<', '>', '<=', '>=', '<>', '+', '-', '*', '/', '\\', '&', '^']);

export function checkStatementForms(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
	memberCtx: MemberCompletionContext = {},
): void {
	checkRemPlacement(source, mod, activity, push);
	// Subs of this module, and of the project's standard modules, by name;
	// a name that is also a Function or a module-level variable anywhere is
	// not judged.
	const subs = new Set<string>();
	const notSubs = new Set<string>();
	for (const symbol of symbols.root.children ?? []) {
		const lower = symbol.name.toLowerCase();
		// A Declare Sub returns nothing either (issue #254).
		if (symbol.kind === 'sub' || (symbol.kind === 'declare' && symbol.declareKind === 'Sub')) {
			subs.add(lower);
		} else {
			notSubs.add(lower);
		}
	}
	for (const [lower, signatures] of projectProcedures ?? []) {
		for (const signature of signatures) {
			(signature.kind === 'sub' ? subs : notSubs).add(lower.toLowerCase());
		}
	}
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const env = typeEnvironmentFor(symbols, member);
		// An object whose default member needs an index: a Collection, or
		// Excel's Hyperlinks, Areas, Borders, Windows, Workbooks and Shapes
		// (issue #221, measured in Excel 16.0).
		const needsIndex = new Map<string, boolean>();
		const indexed = (lower: string): boolean => {
			let answer = needsIndex.get(lower);
			if (answer === undefined) {
				const type = env.get(lower);
				answer = type !== undefined && objectValueNeedsIndex(type, memberCtx);
				needsIndex.set(lower, answer);
			}
			return answer;
		};
		const typedValue = (lower: string): boolean => {
			const type = normalizeType(lower === member.name.toLowerCase() ? member.returnType : env.get(lower));
			return type !== undefined && isKnownScalarType(type);
		};
		const locals = new Set<string>();
		for (const child of procedureSymbolFor(symbols, member)?.children ?? []) {
			locals.add(child.name.toLowerCase());
		}
		forEachStatement(member.body, (stmt) => {
			for (const span of statementAndBranchSpans(stmt)) {
				const toks = statementTokens(source, span);
				const at = (i: number) => ({ start: span.start + toks[i].start, end: span.start + toks[i].end });
				const target = bareAssignmentTarget(source, span);
				// A Set's `=` is the assignment too: `Set c = New Collection` is no
				// operand, and neither is `Set cols(1) = c`, whose target is
				// indexed (issue #140).
				const first = firstExecutableTokenIndex(toks);
				const assigns = target !== undefined || tokenText(toks[first]) === 'set';
				const eq = assigns ? toks.findIndex((tok) => tok.rawText === '=') : -1;
				// A one-line If is judged as its condition here; each branch is its
				// own span with its own assignment (issue #140: `If c Is Nothing
				// Then Set c = New Collection`).
				const then = tokenText(toks[first]) === 'if' && stmt.kind === 'Statement' && stmt.singleLineIfBranches
					? toks.findIndex((tok) => tokenText(tok) === 'then')
					: -1;
				const limit = then > 0 ? then : toks.length;
				for (let i = 0; i < limit; i++) {
					const name = tokenName(toks[i]);
					if (!name || toks[i - 1]?.rawText === '.' || toks[i + 1]?.rawText === ':=' || i === eq - 1) {
						continue;
					}
					// `AddressOf TimerProc` takes the procedure's address, not its value.
					if (tokenText(toks[i - 1]) === 'addressof') {
						continue;
					}
					const lower = name.toLowerCase();
					if (toks[i + 1]?.rawText !== '(' && toks[i + 1]?.rawText !== '.' && indexed(lower)) {
						const typeName = env.get(lower)!;
						const previous = i - 1 === eq ? undefined : toks[i - 1];
						const operator = [toks[i + 1], previous].find((tok) => tok && ((tok.kind === 'operator' && SCALAR_OPERATORS.has(tok.rawText)) || tokenText(tok) === 'mod'));
						if (operator) {
							push('collectionOperand', `'${name}' is ${/^[aeiou]/i.test(typeName) ? 'an' : 'a'} ${typeName}: its default member Item needs an index, so '${operator.rawText}' has no value to work on. This is a VBE compile error: Argument not optional.`, at(i));
							continue;
						}
						// `s = c` with s a String: the whole value of a Let into a
						// typed value (issue #221).
						if (target && i === eq + 1 && toks.length === eq + 2 && typedValue(target.name.toLowerCase())) {
							push('collectionOperand', `'${name}' is ${/^[aeiou]/i.test(typeName) ? 'an' : 'a'} ${typeName}: its default member Item needs an index, so it has no value for '${target.name}' to take. This is a VBE compile error: Argument not optional.`, at(i));
							continue;
						}
					}
					if (target && i > eq && subs.has(lower) && !notSubs.has(lower) && !locals.has(lower) && !env.has(lower)) {
						push('subUsedAsValue', `'${name}' is a Sub, which returns nothing, so it cannot be used as a value. This is a VBE compile error: Expected Function or variable.`, at(i));
					}
				}
			}
		}, activity);
	}
}

const REM_COMMENT = /^rem\b/i;

/**
 * Rules: rem-after-then and rem-after-statement (issues #125 and #231,
 * measured in Excel 16.0). A Rem comment stands at the start of a statement,
 * after a line number or a label, after a block Else, or after a statement
 * in a one-line If's Then or Else list; it may swallow that If's Else. Right
 * after Then, and after any other statement, it is a compile error: "Syntax
 * error" in a procedure, "Expected: end of statement" at module level and on
 * a procedure's own line. The lexer makes it a comment wherever it stands,
 * so its words are never read as code; this judges where it stands.
 */
function checkRemPlacement(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	// Procedure bodies, from the end of the header line to the End line.
	const bodies: Span[] = [];
	for (const member of mod.members) {
		if (member.kind === 'Procedure') {
			const lineEnd = source.indexOf('\n', member.span.start);
			bodies.push({ start: lineEnd < 0 ? member.span.end : lineEnd, end: member.span.end });
		}
	}
	const inBody = (offset: number): boolean => bodies.some((body) => offset > body.start && offset <= body.end);
	let segment: VbaToken[] = [];
	let oneLineIf = false;
	const judge = (endedByColon: boolean): void => {
		const toks = segment;
		segment = [];
		if (toks.length === 0 || toks[0].kind === 'directive') {
			return;
		}
		const last = toks[toks.length - 1];
		const rem = last.kind === 'comment' && REM_COMMENT.test(last.rawText) ? last : undefined;
		const head = isDecimalLineNumber(toks[0]) ? 1 : 0;
		const opener = tokenText(toks[head]);
		const then = opener === 'if' || opener === 'elseif' ? toks.findIndex((tok) => tokenText(tok) === 'then') : -1;
		// The lexer leaves a Rem right after Then a word, so an If stays a
		// one-line If: `If x Then Rem note`, and `ElseIf x Then Rem note`.
		const word = then > head ? toks[then + 1] : undefined;
		if (word && tokenText(word) === 'rem' && !activity?.isInactive(word)) {
			push('remAfterThen', "'Rem' cannot follow 'Then' on one line: a Rem comment starts only at the start of a statement. This is a VBE compile error: Syntax error.", remWord(word));
		}
		if (!oneLineIf && opener === 'if') {
			if (then > head && then < toks.length - 1) {
				oneLineIf = true;
				return;
			}
			// `If x Then:` opens a one-line If too.
			oneLineIf = then > head && endedByColon;
		}
		if (!rem || oneLineIf || toks.length - 1 === head) {
			return;
		}
		if (toks.length - 1 === head + 1 && tokenText(toks[head]) === 'else') {
			return;
		}
		if (activity?.isInactive(rem)) {
			return;
		}
		const error = inBody(rem.start) ? 'Syntax error' : 'Expected: end of statement';
		push('remAfterStatement', `'Rem' starts a comment only at the start of a statement, after a line number, a label or Else, or in a one-line If. Put a colon before it, or use an apostrophe. This is a VBE compile error: ${error}.`, remWord(rem));
	};
	for (const token of tokenizeCached(source)) {
		if (token.kind === 'newline' || token.kind === 'colon') {
			judge(token.kind === 'colon');
			if (token.kind === 'newline') {
				oneLineIf = false;
			}
			continue;
		}
		segment.push(token);
	}
	judge(false);
}

/** The word Rem itself, not the comment it starts. */
function remWord(token: VbaToken): Span {
	return { start: token.start, end: token.start + 3 };
}
