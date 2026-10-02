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
// its Dim, as typeMemberState.ts follows it.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import { resolveRawIntegerConstants, type IntegerConstantLookup } from '../../constants/integerConstantExpression';
import type { HostObjectModel } from '../../host/excelObjectModel';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { ModuleNode, ProcedureNode, Span } from '../../parser/nodes';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { VbaSymbol } from '../../symbols/symbolModel';
import { procedureSymbolFor, type PushFn } from '../analysisContext';
import { collectModuleLiteralIntegerConstants } from '../constExpr';
import { fieldChain, isFixedArrayField, moduleTypes, typeKey, typeRootAt, variableSymbolIn, walkWithSubjects, type FieldStep, type ModuleTypes, type WithSubject } from '../typeFields';
import { isArrayBounds, typeMemberStatesAt, type MemberState, type MemberStatesAt } from '../typeMemberState';
import {
	buildModuleTypeSignatures,
	inferExpressionType,
	isKnownScalarType,
	knownLocalLiteralValuesAt,
	normalizeType,
	procedureIntegerConstantLookup,
	runtimeCallableSourceShadowed,
	sourceNameScopeFor,
	typeEnvironmentFor,
	withKnownLocals,
	type SourceNameScope,
} from '../typeInference';
import { activeModuleMembers, matchParenFrom, pluralizeCount, rawExpressionTokens, statementTokensAfterLeadingLabel, tokenName, tokenText } from '../walker';
import { moduleOptionBase, shapeSubscriptViolation, type FixedArrayBound } from './arrays';
import { isBareOrVbaQualifiedIntrinsicCall } from './shared';

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
	const moduleSignatures = buildModuleTypeSignatures(symbols);
	const optionBase = moduleOptionBase(mod, activity);
	const moduleConstants = collectModuleLiteralIntegerConstants(mod, activity, resolveRawIntegerConstants(projectIntegerConstants ?? new Map(), new Map()));
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const sourceNames = sourceNameScopeFor(symbols, member, projectVisibleSymbols);
		const env = typeEnvironmentFor(symbols, member);
		const valueType = (value: VbaToken[]): string | undefined => normalizeType(inferExpressionType(value, 0, env, moduleSignatures, sourceNames, source)?.type);
		const constants = procedureIntegerConstantLookup(member, moduleConstants, symbols, projectVisibleSymbols, activity, hostModel);
		const valuesAt = knownLocalLiteralValuesAt(source, member, symbols, activity);
		const statesAt: MemberStatesAt | undefined = types.size === 0 ? undefined : typeMemberStatesAt(source, symbols, member, types, activity, optionBase);
		walkWithSubjects(source, member.body, activity, symbols, member, types, undefined, (stmt, subject) => {
			const toks = statementTokensAfterLeadingLabel(source, stmt.span);
			const lookup = withKnownLocals(constants, valuesAt(stmt));
			for (const hit of lenOfArrays(stmt.span, toks, symbols, member, types, subject, sourceNames, valueType)) {
				push(hit.rule, hit.message, hit.span);
			}
			if (!statesAt || ['redim', 'erase'].includes(tokenText(toks[0]))) {
				return; // a ReDim's bounds are not subscripts
			}
			for (let i = 0; i < toks.length; i++) {
				const root = typeRootAt(toks, i, symbols, member, types, subject);
				const states = root ? statesAt(stmt, stmt.span.start + toks[i].start) : undefined;
				const hit = root && states ? chainViolation(stmt.span, toks, i, fieldChain(toks, root, types), states, lookup) : undefined;
				if (hit) {
					push(hit.rule ?? 'arraySubscriptOutOfBounds', hit.message, hit.span);
				}
			}
		});
	}
}

/** The first subscript in a chain that the field cannot take, or a bound it lacks. */
function chainViolation(
	span: Span,
	toks: readonly VbaToken[],
	start: number,
	steps: readonly FieldStep[],
	states: ReadonlyMap<string, MemberState>,
	lookup: IntegerConstantLookup,
): { span: Span; message: string; rule?: 'wrongNumberOfDimensions' } | undefined {
	for (const step of steps) {
		if (!step.field.isArray) {
			continue;
		}
		const state = isFixedArrayField(step.field) || !step.path ? undefined : states.get(step.path);
		const shape: FixedArrayBound | undefined = step.field.dims
			? { name: step.display, dims: step.field.dims, origin: 'Dim' }
			: isArrayBounds(state) ? { ...state, name: step.display } : undefined;
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
	shape: 'unallocated' | FixedArrayBound | undefined,
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
	valueType: (value: VbaToken[]) => string | undefined = () => undefined,
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
		} else if (!array && toks[i - 1]?.rawText !== '.') {
			const type = nonStringValueType(toks.slice(from, to + 1), symbols, proc, valueType);
			if (type) {
				out.push({ span: at, rule: 'variableRequired', message: `${toks[i].rawText} of a ${type} reports a variable's storage size, and '${toks.slice(from, to + 1).map((tok) => tok.rawText).join('')}' is no variable. This is a VBE compile error: Variable required - can't assign to this expression.` });
			}
		} else if (array && array.elementType !== 'byte') {
			out.push({ span: at, rule: 'variantValueMisuse', message: `'${array.display}' is an array, and not of Byte, so VBA.${toks[i].rawText} cannot convert it to a string. This will raise Run-time error '13': Type mismatch.` });
		}
	}
	return out;
}

/** VBA's functions whose return type the library declares as one scalar type: a conversion or Val. */
const TYPED_RETURNS: ReadonlySet<string> = new Set(['cbool', 'cbyte', 'ccur', 'cdate', 'cdbl', 'cint', 'clng', 'clnglng', 'clngptr', 'csng', 'val']);

/**
 * The type of a Len argument that is a value of known type other than
 * String or Variant, and no variable, array element or field: a literal, an
 * expression, a Const, a conversion or Val, or a project Function As Long
 * (issue #368, measured in Excel 16.0). Len(Now), Len(Left(...)) and a
 * Function As Variant or String compile.
 */
function nonStringValueType(
	value: VbaToken[],
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
	valueType: (value: VbaToken[]) => string | undefined,
): string | undefined {
	const name = tokenName(value[0]);
	const named = name ? variableSymbolIn(symbols, proc, name) : undefined;
	const wholeOrIndexed = value.length === 1 || (value[1]?.rawText === '(' && matchParenFrom(value, 1) === value.length - 1);
	if (named && named.kind !== 'constant' && wholeOrIndexed) {
		return undefined;
	}
	const word = value.length === 1 ? tokenText(value[0]) : '';
	if (word === 'true' || word === 'false') {
		return 'Boolean';
	}
	// `Len(i = 1)`: a comparison at the top level is a Boolean.
	let depth = 0;
	for (const tok of value) {
		depth += tok.rawText === '(' ? 1 : tok.rawText === ')' ? -1 : 0;
		if (depth === 0 && (['=', '<>', '<', '>', '<=', '>='].includes(tok.rawText) || ['like', 'is'].includes(tokenText(tok)))) {
			return 'Boolean';
		}
	}
	// `Const K = 5` then `Len(K)`: a Const is no variable.
	const constant = value.length === 1 && name
		? [...(procedureSymbolFor(symbols, proc)?.children ?? []), ...(symbols.root.children ?? [])].find((child) => child.name.toLowerCase() === name.toLowerCase())
		: undefined;
	if (constant?.kind === 'constant') {
		const type = normalizeType(constant.asType) ?? normalizeType(valueType(rawExpressionTokens(constant.defaultRaw ?? '').filter((tok) => tok.kind !== 'comment')));
		return type && type !== 'string' && type !== 'variant' && isKnownScalarType(type) ? type[0].toUpperCase() + type.slice(1) : undefined;
	}
	// A call names a project Function or a typed VBA function; any other
	// library function may return a Variant, as Now does.
	const call = name && value[1]?.rawText === '(' && matchParenFrom(value, 1) === value.length - 1;
	if (call && !TYPED_RETURNS.has(name!.toLowerCase()) && !symbols.root.children?.some((child) => child.kind === 'function' && child.name.toLowerCase() === name!.toLowerCase())) {
		return undefined;
	}
	if (value.length === 1 && name && !named) {
		return undefined;
	}
	const type = valueType(value);
	return type && type !== 'string' && type !== 'variant' && isKnownScalarType(type) ? (type[0].toUpperCase() + type.slice(1)) : undefined;
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
	const root = typeRootAt(toks, from, symbols, proc, types, subject);
	const last = root ? fieldChain(toks, root, types).at(-1) : undefined;
	return last && last.field.isArray && last.open === undefined && last.at === to ? { display: last.display, elementType: last.field.type } : undefined;
}
