// Rule family: the members of a user-defined type (issue #253). Every case
// was measured in Excel 16.0 (build 20326, 2026-10-01).
//
//  - object-variable-not-set: an object member of a Type local that nothing
//    has Set, or that was Set to Nothing, given a member or a call:
//    `t.o.Add 1`, `t.o.Count`, `t.o(1)`, `.o.Add 1` inside `With t`, and
//    `.Add 1` inside `With t.o`. Each raises 91.
//  - scalar-member-access: a member of a member that holds a number or a
//    string, `t.a.Value`, `t.s.Length`, `.a.Value` in `With t`: "Invalid
//    qualifier".
//  - fixed-array-redim: ReDim of a fixed array member, `ReDim t.f(3)`:
//    "Array already dimensioned".
//  - erase-requires-array: Erase of a whole Type value, `Erase t`, or of a
//    member that is no array, `Erase t.a`: "Expected array".
//
// What a member holds comes from typeMemberState.ts.

import type { MemberCompletionContext } from '../../completion/memberAccess';
import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { BodyNode, LeafStatementNode, ModuleNode, ProcedureNode, Span } from '../../parser/nodes';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { PushFn } from '../analysisContext';
import type { DiagnosticRuleName } from '../ruleMetadata';
import { fieldChain, isFixedArrayField, isLeadingDot, moduleTypes, typeKey, typeRootAt, variableSymbolIn, walkWithSubjects, type FieldStep, type ModuleTypes, type WithSubject } from '../typeFields';
import { typeMemberStatesAt, type MemberStatesAt } from '../typeMemberState';
import { isKnownObjectAssignmentType, isKnownScalarType, normalizeType } from '../typeInference';
import { activeModuleMembers, isInactiveNode, matchParenFrom, statementTokens, statementTokensAfterLeadingLabel, tokenName, tokenText } from '../walker';
import { moduleOptionBase } from './arrays';

const NOT_SET = `This will raise Run-time error '91': Object variable or With block variable not set.`;

export function checkTypeMembers(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	memberCtx: MemberCompletionContext,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const types = moduleTypes(source, mod, activity);
	if (types.size === 0) {
		return;
	}
	const optionBase = moduleOptionBase(mod, activity);
	const isObjectType = (type: string): boolean => !types.has(type) && type !== 'variant' && isKnownObjectAssignmentType(type, memberCtx);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const statesAt = typeMemberStatesAt(source, symbols, member, types, activity, optionBase, isObjectType);
		checkForEachFields(source, member.body, symbols, member, types, isObjectType, statesAt, activity, push);
		walkWithSubjects(source, member.body, activity, symbols, member, types, undefined, (stmt, subject) => {
			const toks = statementTokensAfterLeadingLabel(source, stmt.span);
			const head = tokenText(toks[0]);
			if (head === 'redim' || head === 'erase') {
				checkResizes(stmt.span, toks, head, symbols, member, types, subject, isObjectType, push);
				if (head === 'erase') {
					checkEraseOfEmpty(stmt, toks, symbols, member, types, subject, statesAt, push);
				}
				return;
			}
			checkWithObject(stmt, toks, subject, statesAt, push);
			for (let i = 0; i < toks.length; i++) {
				const root = typeRootAt(toks, i, symbols, member, types, subject);
				if (!root) {
					continue;
				}
				const steps = fieldChain(toks, root, types);
				const hit = scalarQualifier(stmt.span, toks, i, steps) ?? nothingAccess(stmt, toks, i, steps, statesAt, isObjectType)
					?? fieldUseMisuse(stmt.span, toks, i, steps, types, isObjectType, symbols, member)
					?? heldValueMisuse(stmt, toks, i, steps, statesAt);
				if (hit) {
					push(hit.rule, hit.message, hit.span);
				}
			}
		});
	}
}

type FieldKind = 'array' | 'scalar' | 'udt' | 'object' | 'variant';

/** What the last field of a chain gives: a whole array, a value of a scalar type, a Type, an object, or a Variant. */
function fieldKind(step: FieldStep, types: ModuleTypes, isObjectType: (type: string) => boolean): FieldKind | undefined {
	if (step.field.isArray && step.open === undefined) {
		return 'array';
	}
	const type = step.field.type;
	const normal = normalizeType(type);
	if (!type || normal === 'variant') {
		return 'variant';
	}
	if (types.has(type)) {
		return 'udt';
	}
	if (isKnownScalarType(normal ?? '')) {
		return 'scalar';
	}
	return isObjectType(type) ? 'object' : undefined;
}

const VALUE_OPERATORS: ReadonlySet<string> = new Set(['+', '-', '*', '/', '\\', '^', 'mod', '&', '=', '<>', '<', '>', '<=', '>=']);

/**
 * A field of a Type used as the shape it is not (issue #417, each measured
 * in Excel 16.0): what the same misuse of a local reports, for `t.f`.
 *
 *  - `t.f(1)` on a number, a string or a Type: "Expected array"
 *    (scalar-indexed). `UBound(t.f)` on what is no array: "Expected array".
 *  - `t.f Is Nothing` on a value, a Type or an array: "Type mismatch".
 *  - `t.f = 5` on an array: "Can't assign to array"; on a Type: "Type
 *    mismatch". `Set t.f = ...` on a Type: "Object required".
 *  - `t.f + 1`, `t.f & "x"` on an array or a Type: "Type mismatch"; on a
 *    Collection, and `Len(t.f)`: "Argument not optional".
 *  - `t.f.Count` on an array: "Invalid qualifier".
 *  - A Type read whole into a Variant, `Main = t.f` or a ByVal Variant
 *    argument: "Only user-defined types defined in public object modules
 *    can be coerced to or from a variant". A Type, an array or a Variant
 *    passed to a ByRef parameter of another type: "ByRef argument type
 *    mismatch".
 */
function fieldUseMisuse(
	span: Span,
	toks: readonly VbaToken[],
	start: number,
	steps: readonly FieldStep[],
	types: ModuleTypes,
	isObjectType: (type: string) => boolean,
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
): { rule: DiagnosticRuleName; message: string; span: Span } | undefined {
	const step = steps.at(-1);
	if (!step) {
		return undefined;
	}
	const end = step.close ?? step.at;
	const at: Span = { start: span.start + toks[start].start, end: span.start + toks[end].end };
	const shown = step.display;
	const declared = step.field.typeName ?? step.field.type ?? 'Variant';
	const after = toks[end + 1];
	const before = toks[start - 1];
	const word = (tok: VbaToken | undefined): string => tokenText(tok);
	// fieldChain stops at the parenthesis after a field that is no array.
	if (!step.field.isArray && toks[step.at + 1]?.rawText === '(') {
		const type = step.field.type;
		if (type && (types.has(type) || (normalizeType(type) !== 'variant' && isKnownScalarType(normalizeType(type) ?? '')))) {
			return { rule: 'scalarIndexed', message: `'${shown}' is declared As ${declared}, which is no array to index. This is a VBE compile error: Expected array.`, span: at };
		}
		return undefined;
	}
	const kind = fieldKind(step, types, isObjectType);
	if (!kind || after?.rawText === '.' && kind !== 'array') {
		return undefined;
	}
	const wholeArgument = (after === undefined || after.rawText === ')' || after.rawText === ',') ;
	if (before?.rawText === '(' && ['ubound', 'lbound'].includes(word(toks[start - 2])) && wholeArgument && kind !== 'array' && kind !== 'variant') {
		return { rule: 'arrayBoundRequiresArray', message: `${toks[start - 2].rawText} takes an array, and '${shown}' is declared As ${declared}. This is a VBE compile error: Expected array.`, span: at };
	}
	if ((word(after) === 'is' || (word(before) === 'is' && word(toks[start - 2]) !== 'typeof')) && (kind === 'scalar' || kind === 'udt' || kind === 'array')) {
		return { rule: 'isOperatorNonObject', message: `The 'Is' operator requires object operands, but '${shown}' is ${kind === 'array' ? 'an array' : `declared As ${declared}`}. This is a VBE compile error: Type mismatch.`, span: at };
	}
	const head = word(toks[0]);
	const target = (start === 0 || (start === 1 && head === 'let')) && after?.rawText === '=';
	const value = target ? toks.slice(end + 2) : [];
	// A dynamic array field takes an array or, As Byte, a String: only a
	// number literal is judged there (issue #417, measured in Excel 16.0).
	const numberValue = value.length === 1 && (value[0].kind === 'integerLiteral' || value[0].kind === 'floatLiteral');
	// A String into one that is not As Byte is refused too (issue #604).
	const elementType = normalizeType(step.field.type);
	const stringValue = value.length === 1 && value[0].kind === 'stringLiteral' && elementType !== 'byte';
	if (target && kind === 'array' && (isFixedArrayField(step.field) || numberValue || stringValue)) {
		return { rule: 'arrayTargetAssignment', message: `'${shown}' is an array, which a value cannot be assigned to whole. This is a VBE compile error: Can't assign to array.`, span: at };
	}
	// `t.f = Split("a b")` gives a String array to a number array: 13
	// (issue #604, measured in Excel 16.0).
	if (target && kind === 'array' && elementType !== undefined && isKnownScalarType(elementType) && !['string', 'byte', 'boolean', 'date'].includes(elementType)
		&& word(value[0]) === 'split' && value[1]?.rawText === '(' && matchParenFrom(value, 1) === value.length - 1) {
		return { rule: 'assignmentTypeMismatch', message: `'${shown}' is an array of ${declared.replace(/\(\)$/, '')}, and Split gives an array of String. This will raise Run-time error '13': Type mismatch.`, span: at };
	}
	if (target && kind === 'udt' && value.length === 1 && ['integerLiteral', 'floatLiteral', 'stringLiteral'].includes(value[0].kind)) {
		return { rule: 'udtValueMismatch', message: `'${shown}' is a ${declared}, which ${value[0].rawText} cannot be assigned to. This is a VBE compile error: Type mismatch.`, span: at };
	}
	if (start === 1 && head === 'set' && after?.rawText === '=' && kind === 'array') {
		return { rule: 'arrayTargetAssignment', message: `'${shown}' is an array, which Set cannot assign to. This is a VBE compile error: Can't assign to array.`, span: at };
	}
	if (start === 1 && head === 'set' && after?.rawText === '=' && kind === 'udt') {
		return { rule: 'setRequiresObject', message: `'${shown}' is a ${declared}, no object variable for Set. This is a VBE compile error: Object required.`, span: at };
	}
	const assignAt = toks.findIndex((tok) => tok.rawText === '=');
	const operatorBefore = before !== undefined && VALUE_OPERATORS.has(word(before)) && start - 1 !== assignAt;
	const operatorAfter = after !== undefined && VALUE_OPERATORS.has(word(after)) && !target && end + 1 !== assignAt;
	if (operatorBefore || operatorAfter) {
		if (kind === 'array' || kind === 'udt') {
			return { rule: 'nonScalarBinaryOperand', message: `The '${(operatorAfter ? after : before)!.rawText}' operator requires a scalar operand, but '${shown}' is ${kind === 'array' ? 'an array' : `a ${declared}`}. This is a VBE compile error: Type mismatch.`, span: at };
		}
		if (kind === 'object' && normalizeType(step.field.type) === 'collection') {
			return { rule: 'collectionOperand', message: `'${shown}' is a Collection: its default member Item needs an index, so '${(operatorAfter ? after : before)!.rawText}' has no value to work on. This is a VBE compile error: Argument not optional.`, span: at };
		}
	}
	if (kind === 'object' && normalizeType(step.field.type) === 'collection' && before?.rawText === '(' && word(toks[start - 2]) === 'len' && after?.rawText === ')') {
		return { rule: 'collectionOperand', message: `'${shown}' is a Collection: its default member Item needs an index, so Len has no value to take. This is a VBE compile error: Argument not optional.`, span: at };
	}
	if (kind === 'array' && after?.rawText === '.' && tokenName(toks[end + 2])) {
		return { rule: 'scalarMemberAccess', message: `Member access on '${shown}' is invalid because it is an array. This is a VBE compile error: Invalid qualifier.`, span: { start: at.start, end: span.start + toks[end + 2].end } };
	}
	// A Type read whole into a Variant: `Main = t.f`.
	if (kind === 'udt' && assignAt === 1 && start === 2 && after === undefined) {
		const lhs = tokenName(toks[0])?.toLowerCase();
		const local = lhs ? variableSymbolIn(symbols, proc, lhs) : undefined;
		const intoVariant = lhs === proc.name.toLowerCase() ? !proc.returnType || normalizeType(proc.returnType) === 'variant' : local !== undefined && !local.isArray && (!local.asType || normalizeType(local.asType) === 'variant');
		if (intoVariant) {
			return { rule: 'udtVariantCoercion', message: `'${shown}' is a ${declared}, which a Variant cannot hold. This is a VBE compile error: Only user-defined types defined in public object modules can be coerced to or from a variant or passed to late-bound functions.`, span: at };
		}
	}
	// An argument of a call statement to a procedure of the module: `TakeV t.f`.
	const callee = tokenName(toks[0])?.toLowerCase();
	const procedure = callee ? (symbols.root.children ?? []).find((child) => (child.kind === 'sub' || child.kind === 'function') && child.name.toLowerCase() === callee) : undefined;
	if (procedure && start >= 1 && (before?.rawText === ',' || start === 1 || (start === 2 && before?.rawText === '(')) && wholeArgument) {
		const position = toks.slice(1, start).filter((tok) => tok.rawText === ',').length;
		const param = (procedure.children ?? []).filter((child) => child.kind === 'parameter')[position];
		const paramType = normalizeType(param?.asType);
		if (param && !param.paramArray && !param.isArray) {
			if (kind === 'udt' && param.byVal && (!paramType || paramType === 'variant')) {
				return { rule: 'udtVariantCoercion', message: `'${shown}' is a ${declared}, which the ByVal Variant '${param.name}' of '${procedure.name}' cannot take. This is a VBE compile error: Only user-defined types defined in public object modules can be coerced to or from a variant or passed to late-bound functions.`, span: at };
			}
			if (!param.byVal && paramType && paramType !== 'variant' && (kind === 'array' || ((kind === 'udt' || kind === 'variant') && paramType !== normalizeType(step.field.type)))) {
				return { rule: 'byRefArgumentTypeMismatch', message: `ByRef argument '${param.name}' of '${procedure.name}' expects ${param.asType}, but '${shown}' is ${kind === 'array' ? 'an array' : `declared As ${declared}`}. This is a VBE compile error: ByRef argument type mismatch.`, span: at };
			}
		}
	}
	return undefined;
}

const ARITHMETIC: ReadonlySet<string> = new Set(['+', '-', '*', '/', '\\', '^', 'mod']);

/**
 * A field read for what it still holds from the Dim (issue #417, each
 * measured in Excel 16.0): an object field nothing has Set, read as a
 * value, raises 91; a Collection Set to a New Collection raises 450 read
 * whole and 5 given an index before anything is added; a String field
 * still "" next to a number in arithmetic raises 13; a Variant field still
 * Empty raises 13 given an index and 424 given a member.
 */
function heldValueMisuse(
	stmt: LeafStatementNode,
	toks: readonly VbaToken[],
	start: number,
	steps: readonly FieldStep[],
	statesAt: MemberStatesAt,
): { rule: DiagnosticRuleName; message: string; span: Span } | undefined {
	const step = steps.at(-1);
	if (!step?.path || step.open !== undefined) {
		return undefined;
	}
	const state = statesAt(stmt, stmt.span.start + toks[start].start).get(step.path);
	if (state !== 'nothing' && state !== 'emptyCollection' && state !== 'emptyString' && state !== 'empty') {
		return undefined;
	}
	const end = step.at;
	const at: Span = { start: stmt.span.start + toks[start].start, end: stmt.span.start + toks[end].end };
	const shown = step.display;
	const head = tokenText(toks[0]);
	const before = toks[start - 1];
	const after = toks[end + 1];
	// The `=` of an assignment statement, not of a comparison.
	const assignAt = toks[0]?.kind === 'identifier' || head === 'let' || head === 'set' ? toks.findIndex((tok) => tok.rawText === '=') : -1;
	if (start + (head === 'let' || head === 'set' ? -1 : 0) === 0 && after?.rawText === '=') {
		return undefined;
	}
	const operatorBefore = before !== undefined && VALUE_OPERATORS.has(tokenText(before)) && start - 1 !== assignAt;
	const operatorAfter = after !== undefined && VALUE_OPERATORS.has(tokenText(after)) && end + 1 !== assignAt;
	const letRead = head !== 'set' && assignAt > 0 && start === assignAt + 1 && after === undefined;
	const call = before?.rawText === '(' && after?.rawText === ')' ? tokenText(toks[start - 2]) : '';
	const isOperand = tokenText(after) === 'is' || (tokenText(before) === 'is' && tokenText(toks[start - 2]) !== 'typeof');
	if (state === 'nothing' && (operatorBefore || operatorAfter || letRead || call === 'len' || call === 'lenb')) {
		return { rule: 'objectVariableNotSet', message: `'${shown}' is an object member that nothing has Set here, so it has no value to read. ${NOT_SET}`, span: at };
	}
	if (state === 'empty' && (call === 'ubound' || call === 'lbound')) {
		return { rule: 'variantValueMisuse', message: `'${shown}' is a Variant that still holds Empty, which is no array for ${toks[start - 2].rawText}. This will raise Run-time error '13': Type mismatch.`, span: at };
	}
	if (state === 'empty' && isOperand) {
		return { rule: 'variantValueMisuse', message: `'${shown}' is a Variant that still holds Empty, not an object, so Is cannot compare it. This will raise Run-time error '424': Object required.`, span: at };
	}
	if (state === 'emptyCollection' && letRead) {
		return { rule: 'objectDefaultValue', message: `'${shown}' is a Collection: its default member Item needs an index, so it has no value to read here. This will raise Run-time error '450': Wrong number of arguments or invalid property assignment.`, span: at };
	}
	if (state === 'emptyCollection' && after?.rawText === '(') {
		return { rule: 'collectionIndexOutOfRange', message: `'${shown}' holds nothing here, so no index reaches an element. This will raise Run-time error '5': Invalid procedure call or argument.`, span: at };
	}
	const numberAt = operatorAfter && ARITHMETIC.has(tokenText(after)) ? end + 2 : operatorBefore && ARITHMETIC.has(tokenText(before)) ? start - 2 : -1;
	const number = toks[numberAt];
	if (state === 'emptyString' && number && (number.kind === 'integerLiteral' || number.kind === 'floatLiteral')) {
		const held = step.field.fixedLength === undefined ? 'holds ""' : 'holds only spaces';
		return { rule: 'stringArithmeticCoercion', message: `Operator '${(numberAt > end ? after : before)!.rawText}' coerces '${shown}', which ${held}, to a number. This will raise Run-time error '13': Type mismatch.`, span: at };
	}
	if (state === 'empty' && after?.rawText === '(') {
		return { rule: 'variantValueMisuse', message: `'${shown}' is a Variant that still holds Empty, which takes no index. This will raise Run-time error '13': Type mismatch.`, span: at };
	}
	if (state === 'empty' && after?.rawText === '.' && tokenName(toks[end + 2])) {
		return { rule: 'variantValueMisuse', message: `'${shown}' is a Variant that still holds Empty, not an object, so it has no ${toks[end + 2].rawText}. This will raise Run-time error '424': Object required.`, span: at };
	}
	return undefined;
}

/** `Erase t.v` while the Variant field still holds Empty: 13 (issue #417, measured in Excel 16.0). */
function checkEraseOfEmpty(
	stmt: LeafStatementNode,
	toks: readonly VbaToken[],
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
	types: ModuleTypes,
	subject: WithSubject | undefined,
	statesAt: MemberStatesAt,
	push: PushFn,
): void {
	for (let i = 1; i < toks.length; i++) {
		if (toks[i - 1].rawText !== ',' && i !== 1) {
			continue;
		}
		const root = typeRootAt(toks, i, symbols, proc, types, subject);
		const step = root ? fieldChain(toks, root, types).at(-1) : undefined;
		const after = step ? toks[(step.close ?? step.at) + 1] : undefined;
		if (!step?.path || step.open !== undefined || (after !== undefined && after.rawText !== ',' && after.kind !== 'comment')) {
			continue;
		}
		if (statesAt(stmt, stmt.span.start + toks[i].start).get(step.path) === 'empty') {
			push('variantValueMisuse', `'${step.display}' is a Variant that still holds Empty, which is not an array for Erase to act on. This will raise Run-time error '13': Type mismatch.`, { start: stmt.span.start + toks[i].start, end: stmt.span.start + toks[step.at].end });
		}
	}
}

/** `For Each e In t.f` with f a number, a string or a Type (issue #417, measured in Excel 16.0). */
function checkForEachFields(
	source: string,
	body: readonly BodyNode[],
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
	types: ModuleTypes,
	isObjectType: (type: string) => boolean,
	statesAt: MemberStatesAt,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const node of body) {
		if (isInactiveNode(activity, node)) {
			continue;
		}
		if (node.kind === 'ForBlock' && node.each && node.sourceExpressionSpan) {
			const span = node.sourceExpressionSpan;
			const toks = statementTokens(source, span);
			const root = typeRootAt(toks, 0, symbols, proc, types, undefined);
			const step = root ? fieldChain(toks, root, types).at(-1) : undefined;
			const kind = step && (step.close ?? step.at) === toks.length - 1 ? fieldKind(step, types, isObjectType) : undefined;
			const state = step?.path && kind ? statesAt(node, span.start).get(step.path) : undefined;
			if (step && (kind === 'scalar' || kind === 'udt')) {
				push('forEachSourceType', `For Each source '${step.display}' must be a collection object or array, but it is declared As ${step.field.typeName}. This is a VBE compile error: For Each may only iterate over a collection object or an array.`, span);
			} else if (step && state === 'nothing') {
				push('objectVariableNotSet', `'${step.display}' is an object member that nothing has Set here, so it is Nothing when For Each asks it for its elements. This will raise Run-time error '424': Object required.`, span);
			} else if (step && state === 'unallocated') {
				push('unallocatedDynamicArrayAccess', `Dynamic array '${step.display}' is not allocated when For Each asks it for its elements. This will raise Run-time error '92': For loop not initialized.`, span);
			} else if (step && state === 'empty') {
				push('variantValueMisuse', `'${step.display}' is a Variant that still holds Empty, which For Each cannot step through. This will raise Run-time error '13': Type mismatch.`, span);
			}
		}
		if ('body' in node && Array.isArray(node.body)) {
			checkForEachFields(source, node.body as BodyNode[], symbols, proc, types, isObjectType, statesAt, activity, push);
		}
	}
}

/** `t.a.Value`: a member of a field that holds a number or a string. */
function scalarQualifier(
	span: Span,
	toks: readonly VbaToken[],
	start: number,
	steps: readonly FieldStep[],
): { rule: 'scalarMemberAccess'; message: string; span: Span } | undefined {
	for (const step of steps) {
		const end = step.close ?? step.at;
		const value = !step.field.isArray || step.open !== undefined;
		const type = normalizeType(step.field.type);
		if (value && type && isKnownScalarType(type) && toks[end + 1]?.rawText === '.') {
			return {
				rule: 'scalarMemberAccess',
				message: `Member access on '${step.display}' is invalid because it is declared as ${step.field.typeName}. This is a VBE compile error: Invalid qualifier.`,
				span: { start: span.start + toks[start].start, end: span.start + toks[end + 1].end },
			};
		}
	}
	return undefined;
}

/** `t.o.Add 1` and `t.o(1)` with t.o still Nothing. */
function nothingAccess(
	stmt: LeafStatementNode,
	toks: readonly VbaToken[],
	start: number,
	steps: readonly FieldStep[],
	statesAt: MemberStatesAt,
	isObjectType: (type: string) => boolean,
): { rule: 'objectVariableNotSet'; message: string; span: Span } | undefined {
	for (const step of steps) {
		const next = toks[step.at + 1]?.rawText;
		if (step.field.isArray || !step.path || !isObjectType(step.field.type ?? '') || (next !== '.' && next !== '!' && next !== '(')) {
			continue;
		}
		if (statesAt(stmt, stmt.span.start + toks[start].start).get(step.path) === 'nothing') {
			return {
				rule: 'objectVariableNotSet',
				message: `'${step.display}' is an object member that nothing has Set here, so it is Nothing. ${NOT_SET}`,
				span: { start: stmt.span.start + toks[start].start, end: stmt.span.start + toks[step.at].end },
			};
		}
		return undefined;
	}
	return undefined;
}

/** `.Add 1` inside `With t.o`, while t.o is Nothing: reported once, at the first member it reaches. */
function checkWithObject(
	stmt: LeafStatementNode,
	toks: readonly VbaToken[],
	subject: WithSubject | undefined,
	statesAt: MemberStatesAt,
	push: PushFn,
): void {
	if (!subject?.path || subject.type || !subject.field || subject.field.isArray) {
		return;
	}
	for (let i = 0; i + 1 < toks.length; i++) {
		if (toks[i].rawText !== '.' || !isLeadingDot(toks, i) || !tokenName(toks[i + 1])) {
			continue;
		}
		if (statesAt(stmt, stmt.span.start + toks[i].start).get(subject.path) === 'nothing') {
			push(
				'objectVariableNotSet',
				`The With object '${subject.display}' is an object member that nothing has Set, so it is Nothing here. ${NOT_SET}`,
				{ start: stmt.span.start + toks[i].start, end: stmt.span.start + toks[i + 1].end },
			);
		}
		return;
	}
}

/** `ReDim t.f(3)` of a fixed member, and `Erase t` or `Erase t.a` of what is no array. */
function checkResizes(
	span: Span,
	toks: readonly VbaToken[],
	head: 'redim' | 'erase',
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
	types: ModuleTypes,
	subject: WithSubject | undefined,
	isObjectType: (type: string) => boolean,
	push: PushFn,
): void {
	const first = head === 'redim' && tokenText(toks[1]) === 'preserve' ? 2 : 1;
	let from = first;
	for (let i = first; i <= toks.length; i++) {
		if (i < toks.length && (toks[i].rawText !== ',' || depthAt(toks, from, i) > 0)) {
			continue;
		}
		const group = toks.slice(from, i).filter((tok) => tok.kind !== 'comment');
		from = i + 1;
		if (group.length === 0) {
			continue;
		}
		const at = (last: VbaToken): Span => ({ start: span.start + group[0].start, end: span.start + last.end });
		if (head === 'erase' && group.length === 1) {
			const variable = variableSymbolIn(symbols, proc, tokenName(group[0])?.toLowerCase() ?? '');
			const type = variable && !variable.isArray ? typeKey(variable.asType) : undefined;
			if (type && types.has(type)) {
				push('eraseRequiresArray', `Erase target '${group[0].rawText}' is a user-defined type, not an array. This is a VBE compile error: Expected array.`, at(group[0]));
			}
			continue;
		}
		const root = typeRootAt(group, 0, symbols, proc, types, subject);
		const step = root ? fieldChain(group, root, types).at(-1) : undefined;
		if (!step) {
			continue;
		}
		// What is no array and no Variant: a number, a string, a Type, a
		// Collection or an Object (issue #417, measured in Excel 16.0).
		const notArray = !step.field.isArray && step.field.type !== undefined && normalizeType(step.field.type) !== 'variant'
			&& (isKnownScalarType(normalizeType(step.field.type) ?? '') || types.has(step.field.type) || isObjectType(step.field.type));
		if (head === 'redim' && isFixedArrayField(step.field) && step.open !== undefined) {
			push('fixedArrayRedim', `'${step.display}' is a fixed-size array member, which ReDim cannot resize. This is a VBE compile error: Array already dimensioned.`, at(group[step.at]));
		} else if (head === 'redim' && notArray && group[step.at + 1]?.rawText === '(') {
			push('scalarRedim', `ReDim target '${step.display}' must be an array or Variant, but it is declared As ${step.field.typeName}. This is a VBE compile error: Expected array.`, at(group[step.at]));
		} else if (head === 'erase' && notArray && step.at === group.length - 1) {
			push('eraseRequiresArray', `Erase target '${step.display}' must be an array or Variant, but it is declared As ${step.field.typeName}. This is a VBE compile error: Expected array.`, at(group[step.at]));
		}
	}
}

/** The parenthesis depth at `index`, counting from `from`. */
function depthAt(toks: readonly VbaToken[], from: number, index: number): number {
	let depth = 0;
	for (let i = from; i < index; i++) {
		depth += toks[i].rawText === '(' ? 1 : toks[i].rawText === ')' ? -1 : 0;
	}
	return depth;
}
