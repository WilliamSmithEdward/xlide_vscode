// Integer constant-expression collection for the diagnostics engine.
//
// Extracted verbatim from `analyzeModule.ts`. The evaluator itself lives in
// `constants/integerConstantExpression.ts` (shared with the structural
// analyzer); this module owns the diagnostics-side collection of raw module
// and procedure-body constants, external (VBA runtime / host) constant
// resolution, and span-based folding of integer expressions.

import type { VbaToken } from '../lexer/tokenKinds';
import {
	enumMemberRawExpression,
	evaluateIntegerConstantExpression,
	parseVbaIntegerLiteral,
	resolveRawIntegerConstants,
	safeInteger,
	type IntegerConstantLookup,
} from '../constants/integerConstantExpression';
import type { ConditionalActivityTracker } from '../conditional/conditionalCompilation';
import type {
	BodyNode,
	EnumNode,
	ModuleNode,
	Span,
	VariableGroupNode,
} from '../parser/nodes';
import type { HostObjectModel } from '../host/excelObjectModel';
import { resolveHostConstant } from '../host/hostModel';
import { resolveRuntimeConstant } from '../runtime/vbaRuntime';
import { activeModuleMembers, isInactiveNode } from './walker';

export function collectModuleLiteralIntegerConstants(
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	base: ReadonlyMap<string, number | undefined> = new Map(),
): Map<string, number | undefined> {
	const rawConstants = new Map<string, string | undefined>();
	const seen = new Set<string>();
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'VariableGroup' && member.isConst) {
			addRawIntegerConstants(member, rawConstants, seen);
		} else if (member.kind === 'Enum') {
			addRawEnumIntegerConstants(member, rawConstants, seen);
		}
	}
	const resolved = new Map(base);
	for (const [name, value] of resolveRawIntegerConstants(rawConstants, base)) {
		resolved.set(name, value);
	}
	return resolved;
}

export function collectBodyLiteralIntegerConstants(
	body: BodyNode[],
	constants: Map<string, number | undefined>,
	activity: ConditionalActivityTracker | undefined,
): void {
	const rawConstants = new Map<string, string | undefined>();
	const seen = new Set<string>();
	collectBodyRawIntegerConstants(body, rawConstants, activity, seen);
	const resolved = resolveRawIntegerConstants(rawConstants, constants);
	for (const [name, value] of resolved) {
		constants.set(name, value);
	}
}

function collectBodyRawIntegerConstants(
	body: BodyNode[],
	rawConstants: Map<string, string | undefined>,
	activity: ConditionalActivityTracker | undefined,
	seen: Set<string>,
): void {
	for (const node of body) {
		if (isInactiveNode(activity, node)) {
			continue;
		}
		if (node.kind === 'VariableGroup') {
			if (node.isConst) {
				addRawIntegerConstants(node, rawConstants, seen);
			}
		} else if ('body' in node && Array.isArray((node as { body?: unknown }).body)) {
			collectBodyRawIntegerConstants((node as { body: BodyNode[] }).body, rawConstants, activity, seen);
		}
	}
}

function addRawIntegerConstants(
	group: VariableGroupNode,
	rawConstants: Map<string, string | undefined>,
	seen: Set<string>,
): void {
	for (const decl of group.declarations) {
		const name = normalizeDeclaredConstantName(decl.name);
		if (!name) {
			continue;
		}
		const key = name.toLowerCase();
		if (seen.has(key)) {
			rawConstants.set(key, undefined);
			continue;
		}
		seen.add(key);
		rawConstants.set(key, decl.defaultRaw);
	}
}

function addRawEnumIntegerConstants(
	en: EnumNode,
	rawConstants: Map<string, string | undefined>,
	seen: Set<string>,
): void {
	let previousName: string | undefined;
	for (const member of en.members) {
		const name = normalizeDeclaredConstantName(member.name);
		if (!name) {
			continue;
		}
		const key = name.toLowerCase();
		if (seen.has(key)) {
			rawConstants.set(key, undefined);
			previousName = name;
			continue;
		}
		seen.add(key);
		rawConstants.set(key, enumMemberRawExpression(member.valueRaw, previousName));
		previousName = name;
	}
}

export function resolveFixedLengthStringSize(
	raw: string,
	constants: ReadonlyMap<string, number | undefined>,
): number | undefined {
	return evaluateIntegerConstantExpression(raw, constants);
}

function normalizeDeclaredConstantName(raw: string): string | undefined {
	const text = raw.trim();
	return text.length > 0 ? text : undefined;
}

/** Qualifiers that name a host's own constant library. Membership is
 * resolved against the CURRENT host's model, so `Word.wdRed` answers in a
 * Word module and misses in an Excel one - Excel-specific names never leak
 * into another host (issue #24). */
const HOST_QUALIFIERS = new Set(['excel', 'word', 'powerpoint', 'access', 'office']);

export function externalIntegerConstantValue(name: string, model?: HostObjectModel): number | undefined {
	const dot = name.indexOf('.');
	if (dot >= 0) {
		const qualifier = name.slice(0, dot);
		const member = name.slice(dot + 1);
		if (qualifier === 'vba') {
			return numericExternalConstantValue(resolveRuntimeConstant(member)?.value);
		}
		if (HOST_QUALIFIERS.has(qualifier)) {
			return numericExternalConstantValue(resolveHostConstant(member, model)?.value);
		}
		return undefined;
	}
	const candidates = [
		numericExternalConstantValue(resolveRuntimeConstant(name)?.value),
		numericExternalConstantValue(resolveHostConstant(name, model)?.value),
	].filter((value): value is number => value !== undefined);
	const unique = [...new Set(candidates)];
	return unique.length === 1 ? unique[0] : undefined;
}

export function numericExternalConstantValue(value: string | number | undefined): number | undefined {
	if (typeof value === 'number') {
		return safeInteger(value);
	}
	if (typeof value !== 'string') {
		return undefined;
	}
	const text = value.trim();
	const signed = /^([+-])(.+)$/.exec(text);
	if (signed) {
		const parsed = parseVbaIntegerLiteral(signed[2]);
		if (parsed === undefined) {
			return undefined;
		}
		return safeInteger(signed[1] === '-' ? -parsed : parsed);
	}
	return parseVbaIntegerLiteral(text);
}

export function foldIntegerExpressionTokens(
	source: string,
	span: Span,
	toks: VbaToken[],
	start: number,
	endExclusive: number,
	constants: IntegerConstantLookup,
): number | undefined {
	if (start >= endExclusive) {
		return undefined;
	}
	const first = toks[start];
	const last = toks[endExclusive - 1];
	const raw = source.slice(span.start + first.start, span.start + last.end);
	return evaluateIntegerConstantExpression(raw, { get: (name) => constants.get(name.toLowerCase()) });
}
