// The user-defined types a module declares, and the fields a dotted path
// reaches through them (issue #248): `t.kids(0).vals` from `Dim t As Outer`.

import type { ConditionalActivityTracker } from '../conditional/conditionalCompilation';
import { evaluateIntegerConstantExpression, type IntegerConstantLookup } from '../constants/integerConstantExpression';
import type { VbaToken } from '../lexer/tokenKinds';
import type { ModuleNode, ProcedureNode } from '../parser/nodes';
import type { buildModuleSymbols } from '../symbols/buildModuleSymbols';
import type { VbaSymbol } from '../symbols/symbolModel';
import { procedureSymbolFor } from './analysisContext';
import { parseFixedArrayBoundsForDecl, type ArrayDimensionBound } from './rules/arrays';
import { activeModuleMembers, matchParenFrom, tokenName } from './walker';

export interface TypeFieldInfo {
	name: string;
	/** The declared type, lowercased; undefined when the field has none. */
	type?: string;
	isArray: boolean;
	/**
	 * A fixed array field's bounds. An implicit lower bound is 0 whatever
	 * Option Base says: `t.arr(0)` runs on `arr(3)` under Option Base 1
	 * (measured in Excel 16.0).
	 */
	dims?: ArrayDimensionBound[];
	/** The raw length of a `String * n` field. */
	fixedLength?: string;
}

/** Each Type the module declares, by lowercased name, to its fields by lowercased name. */
export type ModuleTypes = ReadonlyMap<string, ReadonlyMap<string, TypeFieldInfo>>;

const MODULE_TYPES = new WeakMap<ModuleNode, ModuleTypes>();

/** A declared type's name, lowercased, without the `Module1.` qualifier. */
export function typeKey(asType: string | undefined): string | undefined {
	const name = asType?.trim().split('.').pop()?.trim().toLowerCase();
	return name || undefined;
}

export function moduleTypes(source: string, mod: ModuleNode, activity: ConditionalActivityTracker | undefined): ModuleTypes {
	const cached = MODULE_TYPES.get(mod);
	if (cached) {
		return cached;
	}
	const out = new Map<string, Map<string, TypeFieldInfo>>();
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Type') {
			continue;
		}
		const fields = new Map<string, TypeFieldInfo>();
		for (const field of member.fields) {
			const dims = field.isArray ? parseFixedArrayBoundsForDecl(source, field, 0) : undefined;
			fields.set(field.name.toLowerCase(), {
				name: field.name,
				type: typeKey(field.asType),
				isArray: field.isArray,
				...(dims ? { dims: dims.map((dim) => ({ ...dim, explicitLower: true })) } : {}),
				...(field.fixedLength !== undefined ? { fixedLength: field.fixedLength } : {}),
			});
		}
		out.set(member.name.toLowerCase(), fields);
	}
	MODULE_TYPES.set(mod, out);
	return out;
}

const VARIABLES = new WeakMap<ReturnType<typeof buildModuleSymbols>, Map<ProcedureNode, Map<string, VbaSymbol | undefined>>>();

/** The variable a name means inside a procedure: a parameter or local first, then a module variable. */
export function variableSymbolIn(symbols: ReturnType<typeof buildModuleSymbols>, proc: ProcedureNode, lower: string): VbaSymbol | undefined {
	let byProc = VARIABLES.get(symbols);
	if (!byProc) {
		byProc = new Map();
		VARIABLES.set(symbols, byProc);
	}
	let names = byProc.get(proc);
	if (!names) {
		// Module variables first, so a parameter or local of the same name replaces one.
		names = new Map();
		for (const child of symbols.root.children ?? []) {
			if (child.kind === 'moduleVariable' && !names.has(child.name.toLowerCase())) {
				names.set(child.name.toLowerCase(), child);
			}
		}
		const own = procedureSymbolFor(symbols, proc)?.children ?? [];
		for (const child of [...own].reverse()) {
			names.set(child.name.toLowerCase(), child.kind === 'parameter' || child.kind === 'localVariable' ? child : undefined);
		}
		byProc.set(proc, names);
	}
	return names.get(lower);
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
