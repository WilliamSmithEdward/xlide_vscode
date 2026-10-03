// What a Function of the module returns, where its text fixes it (issue
// #448). A Function that never assigns its name returns its type's default:
// 0, or "" for a String. One whose every assignment is the same literal, and
// that assigns it before anything that may leave, returns that literal as its
// type holds it. So `10 / F()` divides by 0 and `n = F()` stores "abc" in an
// Integer, each measured in Excel 16.0. Anything else that names the result
// (a read, a ByRef pass, a non-literal value), an Err.Raise or Error that may
// end the Function first, or a Variant that is never assigned (Empty) leaves
// the result unknown.

import type { ConditionalActivityTracker } from '../conditional/conditionalCompilation';
import type { VbaToken } from '../lexer/tokenKinds';
import type { BodyNode, ModuleNode, ProcedureNode } from '../parser/nodes';
import { isLeafStatement } from '../parser/nodes';
import type { buildModuleSymbols } from '../symbols/buildModuleSymbols';
import { procedureSymbolFor } from './analysisContext';
import { functionResultFor, normalizeType, stringLiteralValue } from './typeInference';
import {
	activeModuleMembers,
	blockFooterLineSpan,
	blockHeaderLineSpan,
	statementAndBranchSpans,
	rawExpressionTokens,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
} from './walker';

/** A number carries the Function's own type, unless it is a Variant. */
export type FunctionResult = { kind: 'number'; value: number; type?: string } | { kind: 'string'; value: string } | { kind: 'null' };

const INTEGER_TYPES: Readonly<Record<string, readonly [number, number]>> = {
	byte: [0, 255],
	integer: [-32768, 32767],
	long: [-2147483648, 2147483647],
};
const FRACTIONAL_TYPES: ReadonlySet<string> = new Set(['single', 'double', 'currency']);
const SUFFIX_TYPES: Readonly<Record<string, string>> = { '%': 'integer', '&': 'long', '!': 'single', '#': 'double', '@': 'currency', '$': 'string' };
/** Statement heads after which a Function may return before a later assignment. */
const LEAVING_HEADS: ReadonlySet<string> = new Set(['exit', 'goto', 'gosub', 'return', 'end', 'resume', 'on', 'stop', 'error']);

const RESULTS = new WeakMap<ModuleNode, ReadonlyMap<string, FunctionResult>>();

/** Each Function of the module whose result is known, by lowercased name. */
export function knownFunctionResults(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
): ReadonlyMap<string, FunctionResult> {
	let found = RESULTS.get(mod);
	if (found) {
		return found;
	}
	const procedures = new Map<string, ProcedureNode | null>();
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Procedure') {
			const lower = member.name.toLowerCase();
			procedures.set(lower, procedures.has(lower) ? null : member);
		}
	}
	const out = new Map<string, FunctionResult>();
	for (const [lower, proc] of procedures) {
		const result = proc && proc.procKind === 'Function' ? resultOf(source, proc, activity) : undefined;
		if (result) {
			out.set(lower, result);
		}
	}
	found = out;
	RESULTS.set(mod, found);
	CALLS.set(found, { source, procedures, activity });
	return found;
}

/** What a results map needs to run one of its Functions for a call's arguments. */
const CALLS = new WeakMap<ReadonlyMap<string, FunctionResult>, {
	source: string;
	procedures: ReadonlyMap<string, ProcedureNode | null>;
	activity: ConditionalActivityTracker | undefined;
}>();

/**
 * The known result a call stands for at `toks[start]`: `F()`, or a bare `F`
 * not followed by an argument list, with F a Function of the module the
 * calling procedure does not shadow. Undefined otherwise.
 */
export function functionResultAt(
	toks: readonly VbaToken[],
	start: number,
	results: ReadonlyMap<string, FunctionResult>,
	caller: ProcedureNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
): { result: FunctionResult; end: number } | undefined {
	const name = tokenName(toks[start])?.toLowerCase();
	const result = name ? results.get(name) : undefined;
	if (!result || toks[start - 1]?.rawText === '.' || toks[start - 1]?.rawText === '!' || shadowed(name!, caller, symbols)) {
		return undefined;
	}
	if (toks[start + 1]?.rawText === '(') {
		return toks[start + 2]?.rawText === ')' ? { result, end: start + 2 } : undefined;
	}
	return toks[start + 1]?.rawText === '.' || toks[start + 1]?.rawText === '!' ? undefined : { result, end: start };
}

/** The known result of the Function a name calls from `caller`, unless the caller shadows it. */
export function functionResultNamed(
	lower: string,
	results: ReadonlyMap<string, FunctionResult>,
	caller: ProcedureNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
): FunctionResult | undefined {
	const result = results.get(lower);
	return result && !shadowed(lower, caller, symbols) ? result : undefined;
}

/** A whole-number result for an integer lookup's name, `f` or `f()`. */
export function functionIntegerResult(
	name: string,
	results: ReadonlyMap<string, FunctionResult>,
	caller: ProcedureNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
): number | undefined {
	const call = /^(\w+)\((-?\d+(?:,-?\d+)*)?\)$/.exec(name.toLowerCase()) ?? /^(\w+)$/.exec(name.toLowerCase());
	const result = functionResultNamed(name.toLowerCase().replace(/\(\)$/, ''), results, caller, symbols)
		?? (call ? callResult(call[1], call[2] ? call[2].split(',').map(Number) : [], results, caller, symbols) : undefined);
	return result?.kind === 'number' && Number.isInteger(result.value) ? result.value : undefined;
}

/**
 * `Sign1(-1)`: a Function of the module run for whole-number arguments, its
 * result held as its type holds it (issue #562).
 */
function callResult(
	lower: string,
	args: readonly number[],
	results: ReadonlyMap<string, FunctionResult>,
	caller: ProcedureNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
): FunctionResult | undefined {
	const context = CALLS.get(results);
	const proc = context?.procedures.get(lower);
	if (!context || !proc || shadowed(lower, caller, symbols)) {
		return undefined;
	}
	const value = functionResultFor(context.source, proc, symbols, context.activity, args.map((arg) => rawExpressionTokens(String(arg))));
	const literal = value ? literalOf(value.filter((tok) => tok.kind !== 'comment')) : undefined;
	const type = proc.typeSuffix ? SUFFIX_TYPES[proc.typeSuffix] : normalizeType(proc.returnType ?? 'Variant');
	return literal && type ? heldAs(literal, type) : undefined;
}

function shadowed(lower: string, caller: ProcedureNode, symbols: ReturnType<typeof buildModuleSymbols>): boolean {
	return caller.name.toLowerCase() === lower
		|| caller.params.some((param) => param.name.toLowerCase() === lower)
		|| (procedureSymbolFor(symbols, caller)?.children ?? []).some((child) => child.name.toLowerCase() === lower);
}

function resultOf(source: string, proc: ProcedureNode, activity: ConditionalActivityTracker | undefined): FunctionResult | undefined {
	if (/\(/.test(proc.returnType ?? '')) {
		return undefined;
	}
	const type = proc.typeSuffix ? SUFFIX_TYPES[proc.typeSuffix] : normalizeType(proc.returnType ?? 'Variant');
	if (!type) {
		return undefined;
	}
	const lower = proc.name.toLowerCase();
	const literals = new Set<string>();
	let assigned: FunctionResult | undefined;
	let other = false;
	const visit = (toks: readonly VbaToken[]): void => {
		const head = tokenText(toks[0]) === 'let' ? 1 : 0;
		if (tokenText(toks[0]) === 'error' || (tokenText(toks[0]) === 'err' && tokenText(toks[2]) === 'raise')) {
			other = true;
		}
		toks.forEach((tok, i) => {
			if (tokenName(tok)?.toLowerCase() !== lower || toks[i - 1]?.rawText === '.' || toks[i - 1]?.rawText === '!') {
				return;
			}
			const value = i === head && toks[i + 1]?.rawText === '=' ? literalOf(toks.slice(i + 2)) : undefined;
			if (!value) {
				other = true;
				return;
			}
			literals.add(JSON.stringify(value));
			assigned = value;
		});
	};
	walk(source, proc.body, activity, visit);
	if (other || literals.size > 1) {
		return undefined;
	}
	if (!assigned) {
		return type === 'string' ? { kind: 'string', value: '' } : INTEGER_TYPES[type] || FRACTIONAL_TYPES.has(type) || type === 'boolean' ? { kind: 'number', value: 0, type } : undefined;
	}
	const converted = heldAs(assigned, type);
	const held = converted?.kind === 'number' && type !== 'variant' ? { ...converted, type } : converted;
	if (!held) {
		return undefined;
	}
	const isDefault = held.kind === 'number' ? held.value === 0 : held.kind === 'string' ? held.value === '' : false;
	return isDefault || assignsFirst(source, proc.body, activity, lower) ? held : undefined;
}

function walk(source: string, body: readonly BodyNode[], activity: ConditionalActivityTracker | undefined, visit: (toks: readonly VbaToken[]) => void): void {
	for (const node of body) {
		if (activity?.isInactive(node.span) || node.kind === 'ConditionalDirective' || node.kind === 'VariableGroup') {
			continue;
		}
		if (isLeafStatement(node)) {
			for (const span of statementAndBranchSpans(node)) {
				visit(tokens(source, span));
			}
			continue;
		}
		visit(tokens(source, blockHeaderLineSpan(source, node.span)));
		visit(tokens(source, blockFooterLineSpan(source, node.span)));
		if ('body' in node && Array.isArray(node.body)) {
			walk(source, node.body as BodyNode[], activity, visit);
		}
	}
}

/** Whether a top-level statement assigns the result before anything that may leave. */
function assignsFirst(source: string, body: readonly BodyNode[], activity: ConditionalActivityTracker | undefined, lower: string): boolean {
	for (const node of body) {
		if (activity?.isInactive(node.span) || node.kind === 'ConditionalDirective' || node.kind === 'VariableGroup') {
			continue;
		}
		if (isLeafStatement(node)) {
			const toks = tokens(source, node.span);
			const head = tokenText(toks[0]) === 'let' ? 1 : 0;
			if (!(node.kind === 'Statement' && node.singleLineIfBranches) && tokenName(toks[head])?.toLowerCase() === lower && toks[head + 1]?.rawText === '=') {
				return true;
			}
			if (mayLeave(source, [node], activity)) {
				return false;
			}
			continue;
		}
		if (mayLeave(source, [node], activity)) {
			return false;
		}
	}
	return false;
}

function mayLeave(source: string, body: readonly BodyNode[], activity: ConditionalActivityTracker | undefined): boolean {
	let leaves = false;
	walk(source, body, activity, (toks) => {
		if (LEAVING_HEADS.has(tokenText(toks[0])) || (tokenText(toks[0]) === 'err' && tokenText(toks[2]) === 'raise')) {
			leaves = true;
		}
	});
	return leaves;
}

function tokens(source: string, span: { start: number; end: number }): VbaToken[] {
	return statementTokensAfterLeadingLabel(source, span).filter((tok) => tok.kind !== 'comment');
}

function literalOf(toks: readonly VbaToken[]): FunctionResult | undefined {
	if (toks.length === 1 && toks[0].kind === 'stringLiteral') {
		return { kind: 'string', value: stringLiteralValue(toks[0].rawText) };
	}
	const word = toks.length === 1 ? tokenText(toks[0]) : '';
	if (word === 'null') {
		return { kind: 'null' };
	}
	if (word === 'true' || word === 'false') {
		return { kind: 'number', value: word === 'true' ? -1 : 0 };
	}
	const negative = toks.length === 2 && toks[0].rawText === '-';
	const number = toks[negative ? 1 : 0];
	if (toks.length !== (negative ? 2 : 1) || (number.kind !== 'integerLiteral' && number.kind !== 'floatLiteral') || !/^[\d.]+(E[+-]?\d+)?[%&!#@]?$/i.test(number.rawText)) {
		return undefined;
	}
	const value = Number(number.rawText.replace(/[%&!#@]$/, ''));
	return Number.isFinite(value) ? { kind: 'number', value: negative ? -value : value } : undefined;
}

/** A literal as a Function of this type returns it, or undefined where the assignment raises or the value is not plain. */
function heldAs(value: FunctionResult, type: string): FunctionResult | undefined {
	if (type === 'variant') {
		return value;
	}
	if (value.kind !== 'number') {
		return type === 'string' && value.kind === 'string' ? value : undefined;
	}
	const range = INTEGER_TYPES[type];
	if (range) {
		const rounded = roundHalfEven(value.value);
		return rounded >= range[0] && rounded <= range[1] ? { kind: 'number', value: rounded } : undefined;
	}
	if (FRACTIONAL_TYPES.has(type)) {
		return value;
	}
	return type === 'boolean' ? { kind: 'number', value: value.value === 0 ? 0 : -1 } : undefined;
}

function roundHalfEven(value: number): number {
	const floor = Math.floor(value);
	if (value - floor !== 0.5) {
		return Math.round(value);
	}
	return floor % 2 === 0 ? floor : floor + 1;
}
