// Rule: a LongLong or LongPtr narrowed implicitly in 64-bit VBA (issue #298,
// each measured in 64-bit Excel 16.0: a compile error, "Type mismatch").
//
// In 64-bit Office a LongLong, and so a LongPtr, converts to no narrower
// whole-number type on its own: stored in a Long, Integer, Byte or
// Currency, passed ByVal to such a parameter, used as an array index or
// bound, as a For counter's bound, or as a VBA function's whole-number
// argument such as Mid's Start. VarPtr, ObjPtr and StrPtr return a LongPtr,
// so `n = VarPtr(x)`, how 32-bit code keeps a pointer, stops compiling.
// CLng(q) converts; a Double, Variant or String takes the value; `q / 1`
// is a Double. The rule follows the Win64 compiler constant: in 32-bit
// Office LongPtr is a Long and there is no LongLong.

import {
	type ConditionalActivityTracker,
	type ConditionalCompilationEnvironment,
	compilerConstantsWithDefaults,
} from '../../conditional/conditionalCompilation';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { ModuleNode, Span } from '../../parser/nodes';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { VbaProcedureSignature, VbaSymbol } from '../../symbols/symbolModel';
import { procedureSymbolFor, type PushFn } from '../analysisContext';
import { extractCall, isNamedSlot } from '../callExtraction';
import {
	callableTypeSignaturesFor,
	expressionCalls,
	normalizeType,
	runtimeCallableSourceShadowed,
	sourceNameScopeFor,
	typeEnvironmentFor,
} from '../typeInference';
import {
	activeModuleMembers,
	bareAssignmentTarget,
	blockHeaderLineSpan,
	forEachStatement,
	matchParenFrom,
	statementAndBranchSpans,
	statementTokens,
	tokenName,
	tokenText,
} from '../walker';
import { splitTopLevelTokenGroups } from '../../lexer/tokenHelpers';
import { wholeNumberArguments } from './runtimeValues';

const NARROW: ReadonlySet<string> = new Set(['long', 'integer', 'byte', 'currency']);
const WIDE: ReadonlySet<string> = new Set(['longlong', 'longptr']);
const WHOLE: ReadonlySet<string> = new Set(['long', 'integer', 'byte', 'longlong', 'longptr']);
const POINTER_FUNCTIONS: ReadonlySet<string> = new Set(['varptr', 'objptr', 'strptr']);
const OPERATORS: ReadonlySet<string> = new Set(['+', '-', '*', '\\', 'mod']);

export function checkLongLongNarrowing(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	conditionalCompilation: ConditionalCompilationEnvironment | undefined,
	host: string | undefined,
	projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const win64 = compilerConstantsWithDefaults(conditionalCompilation).get('win64');
	if (host?.toLowerCase() === 'vb6' || !(typeof win64 === 'number' ? win64 !== 0 : win64 === true)) {
		return;
	}
	const signatures = callableTypeSignaturesFor(symbols, projectProcedures);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const env = typeEnvironmentFor(symbols, member);
		const sourceNames = sourceNameScopeFor(symbols, member, projectVisibleSymbols);
		const arrays = new Set((procedureSymbolFor(symbols, member)?.children ?? []).filter((child) => child.isArray).map((child) => child.name.toLowerCase()));
		const wide = (toks: readonly VbaToken[]): boolean => wideExpression(toks.filter((tok) => tok.kind !== 'comment'), env, (name) => runtimeCallableSourceShadowed(name, sourceNames));
		const report = (toks: readonly VbaToken[], base: number, where: string): void => {
			const value = toks.filter((tok) => tok.kind !== 'comment');
			push('longLongNarrowing', `${value.map((tok) => tok.rawText).join(' ')} is a LongLong in 64-bit VBA, and ${where} takes no LongLong without a conversion such as CLng. This is a VBE compile error: Type mismatch.`, {
				start: base + value[0].start,
				end: base + value[value.length - 1].end,
			});
		};
		const checkSpan = (span: Span): void => {
			const toks = statementTokens(source, span);
			const target = bareAssignmentTarget(source, span);
			if (target && target.valueTokens.length > 0) {
				const lower = target.name.toLowerCase();
				const type = normalizeType(lower === member.name.toLowerCase() ? member.returnType : env.get(lower));
				if (type && NARROW.has(type) && !arrays.has(lower) && wide(target.valueTokens)) {
					report(target.valueTokens, span.start, `'${target.name}', ${article(type)} ${capitalize(type)},`);
				}
			}
			// `a(q)` and `ReDim a(q)`: an index or a bound.
			toks.forEach((tok, i) => {
				const lower = tokenName(tok)?.toLowerCase();
				if (!lower || !arrays.has(lower) || toks[i + 1]?.rawText !== '(' || toks[i - 1]?.rawText === '.') {
					return;
				}
				const close = matchParenFrom(toks, i + 1);
				for (const slot of close > 0 ? splitTopLevelTokenGroups(toks, i + 2, ',', close) : []) {
					const bounds = splitTopLevelTokenGroups(slot, 0, 'to');
					for (const bound of bounds.length > 0 ? bounds : [slot]) {
						if (bound.length > 0 && wide(bound)) {
							report(bound, span.start, tokenText(toks[0]) === 'redim' ? 'an array bound' : 'an array index');
						}
					}
				}
			});
			// Calls: a ByVal whole-number parameter of a procedure, and a VBA
			// function's whole-number argument.
			const calls = [extractCall(source, span), ...expressionCalls(source, span, signatures, sourceNames)].filter((call) => call !== undefined);
			const seen = new Set<number>();
			for (const call of calls) {
				if (!call || seen.has(call.nameSpan.start) || call.slots.some(isNamedSlot)) {
					continue;
				}
				seen.add(call.nameSpan.start);
				const base = call.sliceStart;
				const signature = call.qualifier ? undefined : signatures.get(call.name.toLowerCase());
				if (signature) {
					call.slots.forEach((slot, k) => {
						const param = signature.params[k];
						const type = normalizeType(param?.type);
						if (param && !param.byRef && type && NARROW.has(type) && slot.length > 0 && wide(slot)) {
							report(slot, base, `the ByVal ${capitalize(type)} '${param.name}' of '${signature.name}'`);
						}
					});
					continue;
				}
			}
			// `Mid$("abc", q, 1)`: a VBA function's whole-number argument, `$` or not.
			toks.forEach((tok, i) => {
				const name = tokenName(tok);
				const open = toks[i + 1]?.rawText === '$' ? i + 2 : i + 1;
				if (!name || toks[open]?.rawText !== '(' || toks[i - 1]?.rawText === '.' || signatures.has(name.toLowerCase()) || runtimeCallableSourceShadowed(name, sourceNames)) {
					return;
				}
				const close = matchParenFrom(toks, open);
				const slots = close > open ? splitTopLevelTokenGroups(toks, open + 1, ',', close) : [];
				for (const slot of wholeNumberArguments(name, slots, host)) {
					if (wide(slot)) {
						report(slot, span.start, `${name}'s whole-number argument`);
					}
				}
			});
		};
		forEachStatement(member.body, (stmt) => {
			for (const span of statementAndBranchSpans(stmt)) {
				checkSpan(span);
			}
		}, activity);
		// A For counter's bounds: `For i = 1 To q` with i a Long.
		const visitFors = (body: typeof member.body): void => {
			for (const node of body) {
				if (activity?.isInactive(node.span) || !('body' in node) || !Array.isArray(node.body)) {
					continue;
				}
				if (node.kind === 'ForBlock' && !node.each && node.controlVariable && NARROW.has(normalizeType(env.get(node.controlVariable.toLowerCase())) ?? '')) {
					const header = blockHeaderLineSpan(source, node.span);
					const toks = statementTokens(source, header);
					const eq = toks.findIndex((tok) => tok.rawText === '=');
					const to = toks.findIndex((tok) => tokenText(tok) === 'to');
					const step = toks.findIndex((tok) => tokenText(tok) === 'step');
					for (const [from, end] of [[eq + 1, to], [to + 1, step > 0 ? step : toks.length], ...(step > 0 ? [[step + 1, toks.length]] : [])]) {
						const part = toks.slice(from, end);
						if (eq > 0 && to > eq && part.length > 0 && wide(part)) {
							report(part, header.start, `the Long counter '${node.controlVariable}'`);
						}
					}
				}
				visitFors(node.body as typeof member.body);
			}
		};
		visitFors(member.body);
	}
}

/**
 * Whether an expression is a whole number that is a LongLong: a LongLong or
 * LongPtr name, a `^` literal or a pointer function, with whole-number
 * operands and +, -, *, \ or Mod between them.
 */
function wideExpression(toks: readonly VbaToken[], env: ReadonlyMap<string, string>, shadowed: (name: string) => boolean): boolean {
	let sawWide = false;
	let i = 0;
	let expectOperand = true;
	while (i < toks.length) {
		const tok = toks[i];
		const word = tokenText(tok);
		if (expectOperand) {
			if (tok.rawText === '-' || tok.rawText === '+') {
				i++;
				continue;
			}
			if (tok.rawText === '(') {
				const close = matchParenFrom(toks, i);
				if (close < 0) {
					return false;
				}
				const inner = toks.slice(i + 1, close);
				const innerWide = wideExpression(inner, env, shadowed);
				if (!innerWide && !wholeExpression(inner, env)) {
					return false;
				}
				sawWide ||= innerWide;
				i = close + 1;
			} else if (tok.kind === 'integerLiteral') {
				sawWide ||= tok.rawText.endsWith('^');
				i++;
			} else if (tokenName(tok) && POINTER_FUNCTIONS.has(word) && toks[i + 1]?.rawText === '(' && toks[i - 1]?.rawText !== '.' && !shadowed(tok.rawText)) {
				const close = matchParenFrom(toks, i + 1);
				if (close < 0) {
					return false;
				}
				sawWide = true;
				i = close + 1;
			} else if (tokenName(tok) && toks[i + 1]?.rawText !== '(' && toks[i + 1]?.rawText !== '.' && toks[i - 1]?.rawText !== '.') {
				const type = normalizeType(env.get(word));
				if (!type || !WHOLE.has(type)) {
					return false;
				}
				sawWide ||= WIDE.has(type);
				i++;
			} else {
				return false;
			}
			expectOperand = false;
			continue;
		}
		if (!OPERATORS.has(word || tok.rawText)) {
			return false;
		}
		expectOperand = true;
		i++;
	}
	return sawWide && !expectOperand;
}

/** Whether an expression is made of whole-number names and literals only. */
function wholeExpression(toks: readonly VbaToken[], env: ReadonlyMap<string, string>): boolean {
	return toks.length > 0 && toks.every((tok) => tok.kind === 'integerLiteral' || OPERATORS.has(tokenText(tok) || tok.rawText)
		|| ['(', ')'].includes(tok.rawText) || WHOLE.has(normalizeType(env.get(tokenText(tok))) ?? ''));
}

function article(type: string): string {
	return /^[aeiou]/i.test(type) ? 'an' : 'a';
}

function capitalize(type: string): string {
	return type.charAt(0).toUpperCase() + type.slice(1);
}
