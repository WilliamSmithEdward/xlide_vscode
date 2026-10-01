// Rule family: arrays one step removed (issue #248). Every case was measured
// in Excel 16.0 (build 20326, 2026-09-30).
//
//  - array-subscript-out-of-bounds: a subscript outside an array field of a
//    user-defined type, `t.arr(5)` on `arr(3) As Long`, at any depth,
//    `t.kids(0).vals(9)`, and inside `With t`; UBound or LBound of a
//    dimension the field lacks; and a dynamic array field nothing has
//    ReDimmed, or that Erase emptied, `t.dyn(0)`, or past the bounds its
//    last ReDim gave it. Each raises 9.
//  - wrong-number-of-dimensions: `t.arr(1, 1)` on a field of one dimension,
//    a compile error.
//  - variable-required: `Len(a)` or `LenB(a)` of an array, a variable or a
//    field, parentheses or not: "Variable required - can't assign to this
//    expression" at compile time.
//  - variant-value-misuse: `VBA.Len(a)` is the library function, which
//    takes the array as a Variant: 13 for any array but a Byte array.
//
// A dynamic field is followed only on a local of the Type, not Static, from
// its Dim: a whole use of the local or of the field (`Fill t`, `t = u`,
// `FillArr t.dyn`), a With, a label and a GoSub end what is known.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import { resolveRawIntegerConstants, type IntegerConstantLookup } from '../../constants/integerConstantExpression';
import type { HostObjectModel } from '../../host/excelObjectModel';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { BodyNode, LeafStatementNode, ModuleNode, ProcedureNode, Span } from '../../parser/nodes';
import { isLeafStatement } from '../../parser/nodes';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { VbaSymbol } from '../../symbols/symbolModel';
import { statementLabelDeclaration } from '../../flow/procedureLabels';
import { procedureSymbolFor, type PushFn } from '../analysisContext';
import { blockHeaderStatements } from '../blockHeaders';
import { collectModuleLiteralIntegerConstants } from '../constExpr';
import { walkEnteringBlocks } from '../dataflow';
import { fieldChain, moduleTypes, typeKey, variableRoot, variableSymbolIn, type FieldStep, type ModuleTypes, type TypeRoot } from '../typeFields';
import {
	knownLocalLiteralValuesAt,
	procedureIntegerConstantLookup,
	runtimeCallableSourceShadowed,
	sourceNameScopeFor,
	withKnownLocals,
	type SourceNameScope,
} from '../typeInference';
import { activeModuleMembers, isInactiveNode, matchParenFrom, pluralizeCount, statementTokensAfterLeadingLabel, tokenName, tokenText } from '../walker';
import { literalDimensions, moduleOptionBase, shapeSubscriptViolation, type FixedArrayBound } from './arrays';
import { isBareOrVbaQualifiedIntrinsicCall } from './shared';

/** A dynamic array field: no elements, or the bounds its last ReDim gave it. */
type FieldState = 'unallocated' | FixedArrayBound;

const SUBSCRIPT_ERROR = `This will raise Run-time error '9': Subscript out of range.`;

export function checkTypeFieldArrays(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
	projectIntegerConstants?: ReadonlyMap<string, string | undefined>,
	projectVisibleSymbols?: readonly VbaSymbol[],
	hostModel?: HostObjectModel,
): void {
	const types = moduleTypes(source, mod, activity);
	const optionBase = moduleOptionBase(mod, activity);
	const moduleConstants = collectModuleLiteralIntegerConstants(mod, activity, resolveRawIntegerConstants(projectIntegerConstants ?? new Map(), new Map()));
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const sourceNames = sourceNameScopeFor(symbols, member, projectVisibleSymbols);
		const constants = procedureIntegerConstantLookup(member, moduleConstants, symbols, projectVisibleSymbols, activity, hostModel);
		const valuesAt = knownLocalLiteralValuesAt(source, member, symbols, activity);
		const statesAt = types.size === 0 ? new Map() : dynamicFieldStatesAt(source, symbols, member, types, activity, optionBase);
		walkWithSubjects(source, member.body, activity, symbols, member, types, undefined, (stmt, subject) => {
			const toks = statementTokensAfterLeadingLabel(source, stmt.span);
			const lookup = withKnownLocals(constants, valuesAt(stmt));
			for (const hit of lenOfArrays(stmt.span, toks, symbols, member, types, subject, sourceNames)) {
				push(hit.rule, hit.message, hit.span);
			}
			if (types.size === 0 || ['redim', 'erase'].includes(tokenText(toks[0]))) {
				return; // a ReDim's bounds are not subscripts
			}
			const states = statesAt.get(stmt) ?? new Map<string, FieldState>();
			for (let i = 0; i < toks.length; i++) {
				const root = rootAt(toks, i, symbols, member, types, subject);
				const hit = root ? chainViolation(stmt.span, toks, i, fieldChain(toks, root, types), states, lookup) : undefined;
				if (hit) {
					push(hit.rule ?? 'arraySubscriptOutOfBounds', hit.message, hit.span);
				}
			}
		});
	}
}

/** The Type value a chain starts from at `i`: a variable, or a leading `.` inside a With. */
function rootAt(
	toks: readonly VbaToken[],
	i: number,
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
	types: ModuleTypes,
	subject: WithSubject | undefined,
): TypeRoot | undefined {
	if (toks[i].rawText === '.') {
		return subject && isLeadingDot(toks, i) ? { ...subject, dot: i } : undefined;
	}
	const lower = tokenName(toks[i])?.toLowerCase();
	if (!lower || toks[i - 1]?.rawText === '.' || toks[i - 1]?.rawText === '!') {
		return undefined;
	}
	return variableRoot(toks, i, variableSymbolIn(symbols, proc, lower), types);
}

/** A `.` that opens an expression, which a With's subject qualifies. */
function isLeadingDot(toks: readonly VbaToken[], i: number): boolean {
	const prev = toks[i - 1];
	if (!prev) {
		return true;
	}
	if (prev.kind === 'keyword') {
		return tokenText(prev) !== 'me';
	}
	return prev.kind === 'operator' || ['(', ',', ':', '='].includes(prev.rawText);
}

/** The first subscript in a chain that the field cannot take, or a bound it lacks. */
function chainViolation(
	span: Span,
	toks: readonly VbaToken[],
	start: number,
	steps: readonly FieldStep[],
	states: ReadonlyMap<string, FieldState>,
	lookup: IntegerConstantLookup,
): { span: Span; message: string; rule?: 'wrongNumberOfDimensions' } | undefined {
	for (const step of steps) {
		if (!step.field.isArray) {
			continue;
		}
		const state = step.field.dims || !step.path ? undefined : states.get(step.path);
		const shape: FixedArrayBound | undefined = step.field.dims
			? { name: step.display, dims: step.field.dims, origin: 'Dim' }
			: state && state !== 'unallocated' ? { ...state, name: step.display } : undefined;
		if (step.open !== undefined) {
			if (state === 'unallocated') {
				return { span: { start: span.start + toks[start].start, end: span.start + toks[step.close!].end }, message: `'${step.display}' ${NO_ELEMENTS}` };
			}
			const hit = shape ? shapeSubscriptViolation(span, toks, shape, step.open, lookup) : undefined;
			if (hit) {
				return hit;
			}
			continue;
		}
		// Used whole: UBound(t.arr, 2), UBound(t.dyn).
		return boundViolation(span, toks, start, step, state === 'unallocated' ? 'unallocated' : shape);
	}
	return undefined;
}

const NO_ELEMENTS = `has no elements here: it is a dynamic array field that nothing has ReDimmed, or that Erase emptied. ${SUBSCRIPT_ERROR}`;

/** `UBound(t.arr, 2)` past the field's dimensions, or UBound of a field with no elements. */
function boundViolation(
	span: Span,
	toks: readonly VbaToken[],
	start: number,
	step: FieldStep,
	shape: FieldState | undefined,
): { span: Span; message: string } | undefined {
	const callee = tokenText(toks[start - 2]);
	const after = toks[step.at + 1]?.rawText;
	if ((callee !== 'ubound' && callee !== 'lbound') || toks[start - 1]?.rawText !== '(' || !isBareOrVbaQualifiedIntrinsicCall(toks, start - 2) || (after !== ')' && after !== ',') || !shape) {
		return undefined;
	}
	const name = toks[start - 2].rawText;
	if (shape === 'unallocated') {
		return { span: { start: span.start + toks[start].start, end: span.start + toks[step.at].end }, message: `${name} reads the bounds of '${step.display}', which ${NO_ELEMENTS}` };
	}
	const close = matchParenFrom(toks, start - 1);
	const dimension = after === ',' && close > step.at + 2 ? toks.slice(step.at + 2, close) : [];
	const value = dimension.length === 1 && dimension[0].kind === 'integerLiteral' ? Number(dimension[0].rawText) : undefined;
	if (value === undefined || (value >= 1 && value <= shape.dims.length)) {
		return undefined;
	}
	return {
		span: { start: span.start + dimension[0].start, end: span.start + dimension[0].end },
		message: `${name} asks for dimension ${value} of '${step.display}', which has ${pluralizeCount(shape.dims.length, 'dimension')}. ${SUBSCRIPT_ERROR}`,
	};
}

/**
 * `Len(a)` and `LenB(a)` of an array: a variable declared as one, or an array
 * field used whole, parentheses or not. `VBA.Len(a)` is the library function
 * instead, which takes the array as a Variant: a Byte array converts to a
 * string and runs, and any other array raises 13 (measured in Excel 16.0).
 */
function lenOfArrays(
	span: Span,
	toks: readonly VbaToken[],
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
	types: ModuleTypes,
	subject: WithSubject | undefined,
	sourceNames: SourceNameScope,
): Array<{ span: Span; message: string; rule: 'variableRequired' | 'variantValueMisuse' }> {
	const out: Array<{ span: Span; message: string; rule: 'variableRequired' | 'variantValueMisuse' }> = [];
	for (let i = 0; i + 1 < toks.length; i++) {
		const word = tokenText(toks[i]);
		if ((word !== 'len' && word !== 'lenb') || toks[i + 1].rawText !== '(' || !isBareOrVbaQualifiedIntrinsicCall(toks, i)) {
			continue;
		}
		if (toks[i - 1]?.rawText !== '.' && runtimeCallableSourceShadowed(toks[i].rawText, sourceNames)) {
			continue;
		}
		const close = matchParenFrom(toks, i + 1);
		let from = i + 2;
		let to = close - 1;
		while (from < to && toks[from].rawText === '(' && matchParenFrom(toks, from) === to) {
			from++;
			to--;
		}
		if (close < 0 || from > to) {
			continue;
		}
		const array = arrayNamed(toks, from, to, symbols, proc, types, subject);
		const at = { start: span.start + toks[from].start, end: span.start + toks[to].end };
		if (array && toks[i - 1]?.rawText !== '.') {
			out.push({ span: at, rule: 'variableRequired', message: `'${array.display}' is an array, which ${toks[i].rawText} cannot take. This is a VBE compile error: Variable required - can't assign to this expression.` });
		} else if (array && array.elementType !== 'byte') {
			out.push({ span: at, rule: 'variantValueMisuse', message: `'${array.display}' is an array, and not of Byte, so VBA.${toks[i].rawText} cannot convert it to a string. This will raise Run-time error '13': Type mismatch.` });
		}
	}
	return out;
}

/** The array the tokens from `from` to `to` name whole, a variable or a field, and its element type, lowercased. */
function arrayNamed(
	toks: readonly VbaToken[],
	from: number,
	to: number,
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
	types: ModuleTypes,
	subject: WithSubject | undefined,
): { display: string; elementType?: string } | undefined {
	if (from === to) {
		const lower = tokenName(toks[from])?.toLowerCase();
		const variable = lower ? variableSymbolIn(symbols, proc, lower) : undefined;
		return variable?.isArray && !variable.paramArray ? { display: toks[from].rawText, elementType: typeKey(variable.asType?.replace(/\(\s*\)\s*$/, '')) } : undefined;
	}
	const root = rootAt(toks, from, symbols, proc, types, subject);
	const last = root ? fieldChain(toks, root, types).at(-1) : undefined;
	return last && last.field.isArray && last.open === undefined && last.at === to ? { display: last.display, elementType: last.field.type } : undefined;
}

/** The Type value a With block's leading `.` stands for. */
type WithSubject = Omit<TypeRoot, 'dot'>;

/**
 * Every statement of a body, block headers included, with the subject of the
 * innermost With around it when that is a value of a module Type.
 */
function walkWithSubjects(
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

/** `With t`, `With t.kid`, `With .kids(0)`: the Type value the header names, or undefined. */
function withSubject(
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
	const root = rootAt(toks, 1, symbols, proc, types, outer);
	const step = root ? fieldChain(toks, root, types).at(-1) : undefined;
	const end = step ? step.close ?? step.at : -1;
	const type = step?.field.type;
	if (!step || end !== last || !type || !types.has(type) || (step.field.isArray && step.open === undefined)) {
		return undefined;
	}
	const display = step.open === undefined ? step.display : step.display + toks.slice(step.open, step.close! + 1).map((tok) => tok.rawText).join('');
	return { type, path: step.open === undefined ? step.path : undefined, display };
}

/**
 * What each dynamic array field of a Type local holds at each statement:
 * nothing from the Dim on, the bounds a literal ReDim gives it, and nothing
 * again after Erase. Blocks are entered as issue #237 enters them.
 */
function dynamicFieldStatesAt(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
	types: ModuleTypes,
	activity: ConditionalActivityTracker | undefined,
	optionBase: number,
): Map<LeafStatementNode, ReadonlyMap<string, FieldState>> {
	const out = new Map<LeafStatementNode, ReadonlyMap<string, FieldState>>();
	const roots = new Map<string, string>();
	let states = new Map<string, FieldState>();
	for (const child of procedureSymbolFor(symbols, proc)?.children ?? []) {
		const type = typeKey(child.asType);
		if (child.kind !== 'localVariable' || child.visibility === 'Static' || child.isArray || !type || !types.has(type)) {
			continue;
		}
		const lower = child.name.toLowerCase();
		roots.set(lower, type);
		for (const key of dynamicFieldPaths(types, type, lower, 0)) {
			states.set(key, 'unallocated');
		}
	}
	if (states.size === 0) {
		return out;
	}
	let seen: ReadonlyMap<string, FieldState> = states;
	const changed = (): void => {
		states = new Map(states);
		seen = states;
	};
	const forget = (names: Iterable<string>): void => {
		for (const name of names) {
			for (const key of [...states.keys()]) {
				if (key === name || key.startsWith(`${name}.`)) {
					if (seen === states) {
						changed();
					}
					states.delete(key);
				}
			}
		}
	};
	const rootsIn = (toks: readonly VbaToken[]): Set<string> => {
		const named = new Set<string>();
		for (let i = 0; i < toks.length; i++) {
			const lower = tokenName(toks[i])?.toLowerCase();
			if (lower && roots.has(lower) && toks[i - 1]?.rawText !== '.' && toks[i - 1]?.rawText !== '!') {
				named.add(lower);
			}
			if (toks[i].rawText === '.' && isLeadingDot(toks, i)) {
				roots.forEach((_, root) => named.add(root));
			}
		}
		return named;
	};
	/** The field path a group names whole, `t.kid.dyn`, and where its parentheses open. */
	const pathOf = (group: readonly VbaToken[]): { path: string; step: FieldStep } | undefined => {
		const lower = tokenName(group[0])?.toLowerCase();
		const type = lower ? roots.get(lower) : undefined;
		const step = type ? fieldChain(group, { type, path: lower, display: group[0].rawText, dot: 1 }, types).at(-1) : undefined;
		return step?.path && states.has(step.path) ? { path: step.path, step } : undefined;
	};
	const visit = (node: BodyNode): void => {
		if (!isLeafStatement(node)) {
			return;
		}
		if (statementLabelDeclaration(source, node.span)) {
			forget(roots.keys());
		}
		const toks = statementTokensAfterLeadingLabel(source, node.span);
		if (states.size > 0) {
			out.set(node, seen);
		}
		const head = tokenText(toks[0]);
		if (head === 'gosub' || (node.kind === 'Statement' && node.singleLineIfBranches !== undefined) || node.singleLineIfTail === true) {
			forget(head === 'gosub' ? roots.keys() : rootsIn(toks));
			return;
		}
		if (head === 'redim' || head === 'erase') {
			const start = tokenText(toks[1]) === 'preserve' ? 2 : 1;
			for (const group of splitGroups(toks.slice(start))) {
				const target = pathOf(group);
				const open = target?.step.open;
				if (target && head === 'erase' && open === undefined && target.step.at === group.length - 1) {
					changed();
					states.set(target.path, 'unallocated');
					continue;
				}
				// A literal ReDim of a field under Option Base 0; Option Base 1 was not measured.
				const dims = target && head === 'redim' && open !== undefined && optionBase === 0 ? literalDimensions(group.slice(open + 1, target.step.close), 0) : undefined;
				if (target && dims) {
					changed();
					states.set(target.path, { name: target.path, dims, origin: 'ReDim' });
					continue;
				}
				forget(target ? [target.path] : rootsIn(group));
			}
			return;
		}
		// A `.` inside a With reaches the With's subject, and entering the block
		// already forgot the root its header names.
		for (let i = 0; i < toks.length; i++) {
			const lower = tokenName(toks[i])?.toLowerCase();
			const type = lower ? roots.get(lower) : undefined;
			if (!lower || !type || toks[i - 1]?.rawText === '.' || toks[i - 1]?.rawText === '!') {
				continue;
			}
			const step = toks[i + 1]?.rawText === '.' ? fieldChain(toks, { type, path: lower, display: toks[i].rawText, dot: i + 1 }, types).at(-1) : undefined;
			if (!step || !step.path) {
				forget(step ? [] : [lower]);
				continue;
			}
			// An element read or write, UBound of the field, or a scalar field leave the bounds alone.
			const bound = ['ubound', 'lbound'].includes(tokenText(toks[i - 2])) && toks[i - 1]?.rawText === '(';
			const isTypeValue = !step.field.isArray && step.field.type !== undefined && types.has(step.field.type);
			if ((step.field.isArray && step.open === undefined && !bound) || isTypeValue) {
				forget([step.path]);
			}
		}
	};
	walkEnteringBlocks(source, proc.body, (node) => isInactiveNode(activity, node), visit, {
		snapshot: () => states,
		restore: (saved) => {
			states = new Map(saved);
			seen = states;
		},
		forget,
		touches: (stmt) => rootsIn(statementTokensAfterLeadingLabel(source, stmt.span)),
	});
	return out;
}

/** `t.dyn`, `t.kid.dyn`: every dynamic array field reached through fields that are not arrays. */
function dynamicFieldPaths(types: ModuleTypes, type: string, prefix: string, depth: number): string[] {
	const out: string[] = [];
	for (const [lower, field] of types.get(type) ?? []) {
		if (field.isArray && !field.dims) {
			out.push(`${prefix}.${lower}`);
		} else if (!field.isArray && field.type && types.has(field.type) && depth < 8) {
			out.push(...dynamicFieldPaths(types, field.type, `${prefix}.${lower}`, depth + 1));
		}
	}
	return out;
}

/** A statement's comma-separated groups outside parentheses. */
function splitGroups(toks: readonly VbaToken[]): VbaToken[][] {
	const out: VbaToken[][] = [[]];
	let depth = 0;
	for (const tok of toks) {
		if (tok.rawText === '(') {
			depth++;
		} else if (tok.rawText === ')') {
			depth--;
		} else if (tok.rawText === ',' && depth === 0) {
			out.push([]);
			continue;
		}
		if (tok.kind !== 'comment') {
			out[out.length - 1].push(tok);
		}
	}
	return out.filter((group) => group.length > 0);
}
