// Rule: an object read as a value when its type has no default member to
// give one (issue #183). Measured in Excel 16.0 (build 20326, 2026-09-29);
// each compiles and raises every time it runs.
//
//  - A class with no default member: `s = c`, `v = c` into a Variant,
//    `Debug.Print c`, `c & "x"`, `c + 1`, `If c = 1` -> 438, Object doesn't
//    support this property or method. So do a Worksheet and a Workbook.
//  - A Collection, whose default member Item needs an index: `v = c` and
//    `Debug.Print c` -> 450, Wrong number of arguments or invalid property
//    assignment. With an operator it does not compile, which is
//    collection-operand's.
//  - An object variable still Nothing raises 91 first.
//
// `Set o = c`, passing `c` to a Variant parameter and `c Is Nothing` read no
// value and run. The write side, `Let c = ...`, is set-required's.

import type { VbaToken } from '../../lexer/tokenKinds';
import type { MemberCompletionContext } from '../../completion/memberAccess';
import type { ProcedureNode, Span } from '../../parser/nodes';
import { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import { procedureSymbolFor, type PushFn } from '../analysisContext';
import { isKnownObjectAssignmentType, normalizeType, objectLetAssignmentVerdict, typeEnvironmentFor } from '../typeInference';
import {
	bareAssignmentTarget,
	firstExecutableTokenIndex,
	statementAndBranchSpans,
	statementTokens,
	tokenName,
	tokenText,
	type ProcedureStatementVisitor,
} from '../walker';

const SCALAR_OPERATORS: ReadonlySet<string> = new Set(['=', '<', '>', '<=', '>=', '<>', '+', '-', '*', '/', '\\', '&', '^']);

export function checkObjectDefaultValues(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	memberCtx: MemberCompletionContext,
	push: PushFn,
): ProcedureStatementVisitor {
	const moduleAutoInstanced = new Set<string>();
	for (const child of symbols.root.children ?? []) {
		if (child.isAutoInstantiated) {
			moduleAutoInstanced.add(child.name.toLowerCase());
		}
	}
	return (proc: ProcedureNode) => {
		const env = typeEnvironmentFor(symbols, proc);
		const verdicts = new Map<string, ReturnType<typeof objectLetAssignmentVerdict>>();
		const verdictFor = (lower: string): ReturnType<typeof objectLetAssignmentVerdict> => {
			let verdict = verdicts.get(lower);
			if (verdict === undefined) {
				// The procedure's own name is its return value only as a target;
				// read, it is a recursive call.
				const type = lower === proc.name.toLowerCase() ? undefined : env.get(lower);
				verdict = type ? objectLetAssignmentVerdict(type, memberCtx) : 'unknown';
				verdicts.set(lower, verdict);
			}
			return verdict;
		};
		// An `As New` variable is never Nothing when read. A local shadows a
		// module-level variable of the same name.
		const autoInstanced = new Set(moduleAutoInstanced);
		for (const child of procedureSymbolFor(symbols, proc)?.children ?? []) {
			if (child.isAutoInstantiated) {
				autoInstanced.add(child.name.toLowerCase());
			} else {
				autoInstanced.delete(child.name.toLowerCase());
			}
		}
		const isObjectVariable = (name: string): boolean => {
			const type = env.get(name.toLowerCase());
			return type !== undefined && isKnownObjectAssignmentType(type, memberCtx);
		};
		return (stmt) => {
			for (const span of statementAndBranchSpans(stmt)) {
				for (const read of valueReads(source, span, stmt.kind === 'Statement' && stmt.singleLineIfBranches !== undefined && span === stmt.span, isObjectVariable)) {
					const lower = tokenName(read.tok)!.toLowerCase();
					const verdict = verdictFor(lower);
					if (verdict !== 'noDefault' && verdict !== 'argument') {
						continue;
					}
					const type = env.get(lower)!;
					// A default member that needs an argument was measured on a
					// Collection only; with an operator it is a compile error,
					// collection-operand's.
					if (verdict === 'argument' && (read.operator || normalizeType(type) !== 'collection')) {
						continue;
					}
					const nothing = autoInstanced.has(lower) ? '' : `, or '91' while it is Nothing`;
					const at: Span = { start: span.start + read.tok.start, end: span.start + read.tok.end };
					push(
						'objectDefaultValue',
						verdict === 'noDefault'
							? `'${read.tok.rawText}' is ${article(type)} ${type}, which has no default member, so it has no value to read here. This will raise Run-time error '438': Object doesn't support this property or method${nothing}.`
							: `'${read.tok.rawText}' is ${article(type)} ${type}: its default member Item needs an index, so it has no value to read here. This will raise Run-time error '450': Wrong number of arguments or invalid property assignment${nothing}.`,
						at,
					);
				}
			}
		};
	};
}

/**
 * The names a statement reads as a value: the whole value of a Let
 * (`s = c`), an item Debug.Print prints (`Debug.Print "a"; c`), and an operand
 * of a scalar operator (`c & "x"`). A name followed by `(` or `.`, or after a
 * `.`, is a call or a member access, not the object's own value.
 */
function valueReads(
	source: string,
	span: Span,
	ifHead: boolean,
	isObjectVariable: (name: string) => boolean,
): Array<{ tok: VbaToken; operator: boolean }> {
	const toks = statementTokens(source, span);
	const first = firstExecutableTokenIndex(toks);
	const head = tokenText(toks[first]);
	if (head === 'set') {
		return [];
	}
	const out: Array<{ tok: VbaToken; operator: boolean }> = [];
	const plainName = (i: number): boolean => tokenName(toks[i]) !== undefined
		&& toks[i - 1]?.rawText !== '.' && toks[i + 1]?.rawText !== '(' && toks[i + 1]?.rawText !== '.';
	const target = bareAssignmentTarget(source, span);
	const eq = target ? toks.findIndex((tok) => tok.rawText === '=') : -1;
	if (target) {
		const value = target.valueTokens.filter((tok) => tok.kind !== 'comment');
		// A Let into an object variable is set-required's.
		if (value.length === 1 && eq + 1 === toks.indexOf(value[0]) && plainName(eq + 1) && !isObjectVariable(target.name)) {
			out.push({ tok: value[0], operator: false });
		}
	}
	if (head === 'debug' && toks[first + 1]?.rawText === '.' && tokenText(toks[first + 2]) === 'print') {
		let depth = 0;
		for (let i = first + 3; i < toks.length; i++) {
			const raw = toks[i].rawText;
			depth += raw === '(' ? 1 : raw === ')' ? -1 : 0;
			const before = toks[i - 1]?.rawText;
			const after = toks[i + 1]?.rawText;
			const alone = (i === first + 3 || before === ',' || before === ';') && (after === undefined || after === ',' || after === ';' || toks[i + 1].kind === 'comment');
			if (depth === 0 && alone && plainName(i)) {
				out.push({ tok: toks[i], operator: false });
			}
		}
	}
	const then = ifHead ? toks.findIndex((tok) => tokenText(tok) === 'then') : -1;
	const limit = then > 0 ? then : toks.length;
	for (let i = first; i < limit; i++) {
		if (i === eq - 1 || !plainName(i) || out.some((read) => read.tok === toks[i])) {
			continue;
		}
		const previous = i - 1 === eq ? undefined : toks[i - 1];
		const isOperator = (tok: VbaToken | undefined): boolean => tok !== undefined
			&& ((tok.kind === 'operator' && SCALAR_OPERATORS.has(tok.rawText)) || tokenText(tok) === 'mod');
		if (isOperator(toks[i + 1]) || isOperator(previous)) {
			out.push({ tok: toks[i], operator: true });
		}
	}
	return out;
}

function article(type: string): string {
	return /^[aeiou]/i.test(type) ? 'an' : 'a';
}
