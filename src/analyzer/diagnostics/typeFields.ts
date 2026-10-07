// The user-defined types a module declares, and the fields a dotted path
// reaches through them (issue #248): `t.kids(0).vals` from `Dim t As Outer`.

import type { ConditionalActivityTracker } from '../conditional/conditionalCompilation';
import { evaluateIntegerConstantExpression, type IntegerConstantLookup } from '../constants/integerConstantExpression';
import type { VbaToken } from '../lexer/tokenKinds';
import type { BodyNode, LeafStatementNode, ModuleNode, ProcedureNode } from '../parser/nodes';
import { isLeafStatement } from '../parser/nodes';
import type { buildModuleSymbols } from '../symbols/buildModuleSymbols';
import type { VbaSymbol } from '../symbols/symbolModel';
import { procedureSymbolFor } from './analysisContext';
import { blockHeaderStatements } from './blockHeaders';
import { parseFixedArrayBoundsForDecl, type ArrayDimensionBound } from './rules/arrays';
import { collectModuleLiteralIntegerConstants } from './constExpr';
import { splitTopLevelTokenGroups, statementTokensCached } from '../lexer/tokenHelpers';
import { activeModuleMembers, isInactiveNode, matchParenFrom, statementTokensAfterLeadingLabel, tokenName, tokenText } from './walker';

export interface TypeFieldInfo {
	name: string;
	/** The declared type, lowercased; undefined when the field has none. */
	type?: string;
	/** The declared type as written, for a message. */
	typeName?: string;
	isArray: boolean;
	/**
	 * A fixed array field's bounds. An implicit lower bound is 0 whatever
	 * Option Base says: `t.arr(0)` runs on `arr(3)` under Option Base 1
	 * (measured in Excel 16.0).
	 */
	dims?: ArrayDimensionBound[];
	/**
	 * A fixed array field whose bounds are not known here: `vals(1 To N)`
	 * with N a Public Const of another module (issue #366).
	 */
	sized?: boolean;
	/** The raw length of a `String * n` field. */
	fixedLength?: string;
}

/** Each Type the module declares, by lowercased name, to its fields by lowercased name. */
export type ModuleTypes = ReadonlyMap<string, ReadonlyMap<string, TypeFieldInfo>>;

const MODULE_TYPES = new WeakMap<ModuleNode, {
	source: string;
	activity: ConditionalActivityTracker | undefined;
	result: ModuleTypes;
}>();

/** A declared type's name, lowercased, without the `Module1.` qualifier. */
export function typeKey(asType: string | undefined): string | undefined {
	const name = asType?.trim().split('.').pop()?.trim().toLowerCase();
	return name || undefined;
}

export function moduleTypes(source: string, mod: ModuleNode, activity: ConditionalActivityTracker | undefined): ModuleTypes {
	const cached = MODULE_TYPES.get(mod);
	if (cached && cached.source === source && cached.activity === activity) {
		return cached.result;
	}
	const out = new Map<string, Map<string, TypeFieldInfo>>();
	let constants: IntegerConstantLookup | undefined;
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Type') {
			continue;
		}
		const fields = new Map<string, TypeFieldInfo>();
		for (const field of member.fields) {
			// Bounds a module Const or Enum member states are read too, and
			// bounds that cannot be read still make the field fixed (issue #366).
			const bounds = field.isArray ? boundsTokens(source, field.span) : undefined;
			const dims = !field.isArray || !bounds?.length ? undefined
				: parseFixedArrayBoundsForDecl(source, field, 0) ?? constantDimensions(bounds, constants ??= collectModuleLiteralIntegerConstants(mod, activity));
			const sized = !dims && (bounds?.length ?? 0) > 0;
			fields.set(field.name.toLowerCase(), {
				name: field.name,
				type: typeKey(field.asType),
				typeName: field.asType?.trim(),
				isArray: field.isArray,
				...(dims ? { dims: dims.map((dim) => ({ ...dim, explicitLower: true })) } : {}),
				...(sized ? { sized } : {}),
				...(field.fixedLength !== undefined ? { fixedLength: field.fixedLength } : {}),
			});
		}
		out.set(member.name.toLowerCase(), fields);
	}
	MODULE_TYPES.set(mod, { source, activity, result: out });
	return out;
}

/** Whether an array field is fixed, its bounds known or not. */
export function isFixedArrayField(field: TypeFieldInfo): boolean {
	return field.isArray && (field.dims !== undefined || field.sized === true);
}

/** The tokens between a declaration's parentheses, comments dropped; undefined without them. */
function boundsTokens(source: string, span: { start: number; end: number }): VbaToken[] | undefined {
	const toks = statementTokensCached(source, span);
	const open = toks.findIndex((tok) => tok.rawText === '(');
	const close = open < 0 ? -1 : matchParenFrom(toks, open);
	return close < 0 ? undefined : toks.slice(open + 1, close).filter((tok) => tok.kind !== 'comment');
}

/** The bounds `1 To N, N * 2` state with every name a module Const or Enum member. An implicit lower bound is 0. */
function constantDimensions(bounds: readonly VbaToken[], constants: IntegerConstantLookup): ArrayDimensionBound[] | undefined {
	const out: ArrayDimensionBound[] = [];
	for (const dim of splitTopLevelTokenGroups(bounds, 0, ',')) {
		const to = dim.findIndex((tok) => tokenText(tok) === 'to');
		const text = (part: readonly VbaToken[]): string => part.map((tok) => tok.rawText).join(' ');
		const lower = to < 0 ? 0 : evaluateIntegerConstantExpression(text(dim.slice(0, to)), constants);
		const upper = evaluateIntegerConstantExpression(text(to < 0 ? dim : dim.slice(to + 1)), constants);
		if (lower === undefined || upper === undefined || dim.length === 0) {
			return undefined;
		}
		out.push({ lower, upper, explicitLower: true });
	}
	return out.length > 0 ? out : undefined;
}

const VARIABLES = new WeakMap<ReturnType<typeof buildModuleSymbols>, {
	module: ReadonlyMap<string, VbaSymbol>;
	procedures: WeakMap<ProcedureNode, ReadonlyMap<string, VbaSymbol | undefined>>;
}>();

/** The variable a name means inside a procedure: a parameter or local first, then a module variable. */
export function variableSymbolIn(symbols: ReturnType<typeof buildModuleSymbols>, proc: ProcedureNode, lower: string): VbaSymbol | undefined {
	let scopes = VARIABLES.get(symbols);
	if (!scopes) {
		const module = new Map<string, VbaSymbol>();
		for (const child of symbols.root.children ?? []) {
			if (child.kind === 'moduleVariable' && !module.has(child.name.toLowerCase())) {
				module.set(child.name.toLowerCase(), child);
			}
		}
		scopes = { module, procedures: new WeakMap() };
		VARIABLES.set(symbols, scopes);
	}
	let names = scopes.procedures.get(proc);
	if (!names) {
		// Keep only the procedure overlay; an explicit undefined shadows a
		// module variable with a declaration that is not a variable.
		const ownNames = new Map<string, VbaSymbol | undefined>();
		const own = procedureSymbolFor(symbols, proc)?.children ?? [];
		for (const child of [...own].reverse()) {
			ownNames.set(child.name.toLowerCase(), child.kind === 'parameter' || child.kind === 'localVariable' ? child : undefined);
		}
		names = ownNames;
		scopes.procedures.set(proc, names);
	}
	return names.has(lower) ? names.get(lower) : scopes.module.get(lower);
}

/** A value of a module Type that a chain of fields starts from. */
export interface TypeRoot {
	type: string;
	/** The variable and fields so far, lowercased, while no subscript intervenes. */
	path?: string;
	/** The text so far, as written: `ts(1)`, `t`. */
	display: string;
	/** Index of the `.` that follows the root. */
	dot: number;
}

/**
 * The Type value a variable at `i` is, when a `.` follows it: `t.` with `Dim
 * t As Outer`, or `ts(1).` with `Dim ts(2) As Outer`.
 */
export function variableRoot(toks: readonly VbaToken[], i: number, variable: VbaSymbol | undefined, types: ModuleTypes): TypeRoot | undefined {
	const type = variable ? typeKey(variable.asType) : undefined;
	if (!variable || !type || !types.has(type) || variable.fixedLength !== undefined) {
		return undefined;
	}
	if (!variable.isArray) {
		return toks[i + 1]?.rawText === '.' ? { type, path: variable.name.toLowerCase(), display: toks[i].rawText, dot: i + 1 } : undefined;
	}
	const close = toks[i + 1]?.rawText === '(' ? matchParenFrom(toks, i + 1) : -1;
	return close > 0 && toks[close + 1]?.rawText === '.'
		? { type, display: toks.slice(i, close + 1).map((tok) => tok.rawText).join(''), dot: close + 1 }
		: undefined;
}

/**
 * The declared length of the fixed-length string the tokens name: `s` with
 * `Dim s As String * 3`, an element `s(1)` of an array of them, or a field
 * `t.name` declared `name As String * 3`. A Const length is read too.
 */
export function fixedStringLength(
	toks: readonly VbaToken[],
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
	types: ModuleTypes,
	constants: IntegerConstantLookup,
): number | undefined {
	const lower = tokenName(toks[0])?.toLowerCase();
	const variable = lower ? variableSymbolIn(symbols, proc, lower) : undefined;
	let raw: string | undefined;
	if (variable?.fixedLength !== undefined) {
		const element = toks[1]?.rawText === '(' && matchParenFrom(toks, 1) === toks.length - 1;
		raw = (variable.isArray ? element : toks.length === 1) ? variable.fixedLength : undefined;
	} else {
		const root = variableRoot(toks, 0, variable, types);
		const last = root ? fieldChain(toks, root, types).at(-1) : undefined;
		const whole = last !== undefined && (last.close ?? last.at) === toks.length - 1 && (!last.field.isArray || last.open !== undefined);
		raw = whole ? last.field.fixedLength : undefined;
	}
	return raw === undefined ? undefined : evaluateIntegerConstantExpression(raw, constants);
}

/** A `.` that opens an expression, which a With's subject qualifies. */
export function isLeadingDot(toks: readonly VbaToken[], i: number): boolean {
	const prev = toks[i - 1];
	if (!prev) {
		return true;
	}
	if (prev.kind === 'keyword') {
		return tokenText(prev) !== 'me';
	}
	// `FillArr .dyn`: a space after a name opens a call's argument.
	if ((prev.kind === 'identifier' || prev.kind === 'bracketedIdentifier') && prev.end < toks[i].start) {
		return true;
	}
	return prev.kind === 'operator' || ['(', ',', ':', '='].includes(prev.rawText);
}

/** What a With block's leading `.` stands for. */
export interface WithSubject {
	/** The subject as written: `t`, `t.kids(0)`. */
	display: string;
	/** The variable and fields, lowercased, while no subscript intervenes. */
	path?: string;
	/** The module Type the subject is a value of. */
	type?: string;
	/** The field the subject ends at, when it is one: `With t.o` is the object field o. */
	field?: TypeFieldInfo;
}

/** The Type value a chain starts from at `i`: a variable, or a leading `.` inside a With of a Type value. */
export function typeRootAt(
	toks: readonly VbaToken[],
	i: number,
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
	types: ModuleTypes,
	subject: WithSubject | undefined,
): TypeRoot | undefined {
	if (toks[i].rawText === '.') {
		return subject?.type && isLeadingDot(toks, i) ? { type: subject.type, path: subject.path, display: subject.display, dot: i } : undefined;
	}
	const lower = tokenName(toks[i])?.toLowerCase();
	if (!lower || toks[i - 1]?.rawText === '.' || toks[i - 1]?.rawText === '!') {
		return undefined;
	}
	return variableRoot(toks, i, variableSymbolIn(symbols, proc, lower), types);
}

/**
 * Every statement of a body, block headers included, with the subject of the
 * innermost With around it when that is a value of a module Type or a field of
 * one.
 */
export function walkWithSubjects(
	source: string,
	body: readonly BodyNode[],
	activity: ConditionalActivityTracker | undefined,
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
	types: ModuleTypes,
	subject: WithSubject | undefined,
	visit: (stmt: LeafStatementNode, subject: WithSubject | undefined) => void,
): void {
	for (const node of body) {
		if (isInactiveNode(activity, node)) {
			continue;
		}
		if (isLeafStatement(node)) {
			visit(node, subject);
			continue;
		}
		if (!('body' in node) || !Array.isArray(node.body)) {
			continue;
		}
		const { before, after } = blockHeaderStatements(source, node);
		if (before) {
			visit(before, subject);
		}
		const inner = node.kind === 'WithBlock'
			? (before ? withSubject(statementTokensAfterLeadingLabel(source, before.span), symbols, proc, types, subject) : undefined)
			: subject;
		walkWithSubjects(source, node.body, activity, symbols, proc, types, inner, visit);
		if (after) {
			visit(after, subject);
		}
	}
}

/** `With t`, `With t.kid`, `With .kids(0)`, `With t.o`: what the header names, or undefined. */
export function withSubject(
	toks: readonly VbaToken[],
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
	types: ModuleTypes,
	outer: WithSubject | undefined,
): WithSubject | undefined {
	if (tokenText(toks[0]) !== 'with' || toks.length < 2) {
		return undefined;
	}
	const last = toks.length - 1;
	if (last === 1) {
		const variable = variableSymbolIn(symbols, proc, tokenName(toks[1])?.toLowerCase() ?? '');
		const type = variable && !variable.isArray ? typeKey(variable.asType) : undefined;
		return type && types.has(type) ? { type, path: variable!.name.toLowerCase(), display: toks[1].rawText } : undefined;
	}
	const root = typeRootAt(toks, 1, symbols, proc, types, outer);
	const step = root ? fieldChain(toks, root, types).at(-1) : undefined;
	const end = step ? step.close ?? step.at : -1;
	if (!step || end !== last || (step.field.isArray && step.open === undefined)) {
		return undefined;
	}
	const type = step.field.type !== undefined && types.has(step.field.type) ? step.field.type : undefined;
	const display = step.open === undefined ? step.display : step.display + toks.slice(step.open, step.close! + 1).map((tok) => tok.rawText).join('');
	return { display, path: step.open === undefined ? step.path : undefined, type, field: step.field };
}

/** The subject of the innermost With around each statement, block headers included, by the statement's start. */
export function withSubjectsIn(
	source: string,
	proc: ProcedureNode,
	activity: ConditionalActivityTracker | undefined,
	symbols: ReturnType<typeof buildModuleSymbols>,
	types: ModuleTypes,
): Map<number, WithSubject> {
	const out = new Map<number, WithSubject>();
	walkWithSubjects(source, proc.body, activity, symbols, proc, types, undefined, (stmt, subject) => {
		if (subject) {
			out.set(stmt.span.start, subject);
		}
	});
	return out;
}

/** One field a chain reaches. */
export interface FieldStep {
	field: TypeFieldInfo;
	/** `t.kid.dyn`, lowercased; undefined once a subscript intervenes. */
	path?: string;
	/** `t.kids(0).vals`, as written. */
	display: string;
	/** Index of the field's name. */
	at: number;
	/** The parentheses right after the field, when it is indexed. */
	open?: number;
	close?: number;
}

/**
 * The fields a chain reaches from a Type value whose `.` is at `root.dot`:
 * `t.kids(0).vals(9)` reaches kids, then vals. It stops at a name that is
 * no field of a module Type, and at an array field used whole.
 */
export function fieldChain(toks: readonly VbaToken[], root: TypeRoot, types: ModuleTypes): FieldStep[] {
	const steps: FieldStep[] = [];
	let type: string | undefined = root.type;
	let path = root.path;
	let display = root.display;
	let i = root.dot;
	while (type && toks[i]?.rawText === '.') {
		const lower = tokenName(toks[i + 1])?.toLowerCase();
		const field: TypeFieldInfo | undefined = lower ? types.get(type)?.get(lower) : undefined;
		if (!field || !lower) {
			break;
		}
		path = path === undefined ? undefined : `${path}.${lower}`;
		display = `${display}.${toks[i + 1].rawText}`;
		const step: FieldStep = { field, path, display, at: i + 1 };
		steps.push(step);
		i += 2;
		if (field.isArray) {
			const close = toks[i]?.rawText === '(' ? matchParenFrom(toks, i) : -1;
			if (close < 0) {
				break;
			}
			step.open = i;
			step.close = close;
			path = undefined;
			display += toks.slice(i, close + 1).map((tok) => tok.rawText).join('');
			i = close + 1;
		} else if (toks[i]?.rawText === '(') {
			break;
		}
		type = field.type !== undefined && types.has(field.type) ? field.type : undefined;
	}
	return steps;
}
