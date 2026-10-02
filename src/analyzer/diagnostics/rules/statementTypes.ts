// Rule family: statements the VBE refuses by their form or by the type of
// what they name (issue #213); statementForms.ts holds the #125 family.
// Each verdict is a full project compile in 64-bit Excel 16.0 (build 20326,
// 2026-09-30):
//
//  - for-variable-in-use: a For or For Each inside another on the same
//    control variable, "For control variable already in use".
//  - for-counter-type: a For counter that is a String, Boolean, object or
//    user-defined type, "Type mismatch". Variant, Date and the numbers take.
//  - for-each-source-type: For Each over an array of a user-defined type or
//    of fixed-length strings, fixed or dynamic.
//  - statement-before-first-case: a statement or label between Select Case
//    and its first Case; a comment is fine.
//  - return-with-value: `Return 5`, "Syntax error". A bare Return is GoSub's.
//  - with-scalar-target: With on a number, a string or a scalar variable,
//    "With object must be user-defined type, Object, or Variant".
//  - set-requires-object: Set of a user-defined type variable.
//  - paramarray-passing-mode: ByVal or ByRef on a ParamArray, "Expected:
//    identifier".
//  - udt-value-mismatch and udt-variant-coercion: a user-defined type where
//    a single value is needed, and one handed to a Variant.
//
// Issue #216, measured the same way:
//
//  - const-value-not-constant and array-bound-not-constant: a Const value or
//    a Dim bound that names a variable, "Constant expression required".
//    ReDim takes one.
//  - variable-required: a Const as a For counter or a Mid target, "Variable
//    required - can't assign to this expression".
//  - type-suffix-mismatch: `n% = 2` with n As Long, "Type-declaration
//    character does not match declared data type". The suffix of the
//    declared type is fine.
//  - named-argument-not-allowed: InStr, Len, StrComp, Abs, Int, Fix, Sgn and
//    the C-conversions but CDec take no named arguments, "Syntax error".
//
// Issue #253, measured the same way:
//
//  - lset-type-mismatch: LSet between two different user-defined types,
//    either of which holds a variable-length String, a dynamic array, an
//    object or a Variant, "Type mismatch". Two Types of fixed-size
//    members, and two values of one Type, take it.
//  - udt-variant-coercion: a user-defined type passed to a method of a
//    Collection, an Object or a Variant, `c.Add t`, `o.Add t`.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { BodyNode, ModuleNode, ProcedureNode, Span } from '../../parser/nodes';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { VbaSymbol } from '../../symbols/symbolModel';
import type { ProjectTypeName } from '../../completion/typeCompletion';
import { procedureSymbolFor, type PushFn } from '../analysisContext';
import { isKnownScalarType, normalizeType, sourceIdentifierBinding } from '../typeInference';
import { isFixedArrayField, moduleTypes, type ModuleTypes } from '../typeFields';
import {
	absoluteSpan,
	activeModuleMembers,
	isInactiveNode,
	matchParenFrom,
	rawExpressionTokens,
	statementTokens,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
} from '../walker';

/** What a variable holds, as far as these rules care. */
type Category = 'variant' | 'number' | 'date' | 'string' | 'boolean' | 'object' | 'udt' | 'class';

const NUMBER_TYPES = new Set(['byte', 'integer', 'long', 'longlong', 'longptr', 'currency', 'single', 'double', 'decimal']);

/** Words that join two operands, so a user-defined type beside one is used as a value. */
const OPERATOR_WORDS = new Set(['and', 'or', 'xor', 'eqv', 'imp', 'mod', 'like', 'not', 'is']);

interface Context {
	source: string;
	symbols: ReturnType<typeof buildModuleSymbols>;
	projectVisibleSymbols: readonly VbaSymbol[] | undefined;
	activity: ConditionalActivityTracker | undefined;
	push: PushFn;
	udts: ReadonlySet<string>;
	enums: ReadonlySet<string>;
	/** Whether an untyped variable is a Variant, which a Deftype line changes. */
	untypedIsVariant: boolean;
	/** The Types this module declares, with their fields. */
	types: ModuleTypes;
}

export function checkStatementTypes(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	projectTypes: readonly ProjectTypeName[] | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const udts = new Set<string>();
	const enums = new Set<string>();
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Type') {
			udts.add(member.name.toLowerCase());
		} else if (member.kind === 'Enum') {
			enums.add(member.name.toLowerCase());
		}
	}
	for (const type of projectTypes ?? []) {
		if (type.kind === 'userType') {
			udts.add(type.name.toLowerCase());
		} else if (type.kind === 'enum') {
			enums.add(type.name.toLowerCase());
		}
	}
	const ctx: Context = {
		source,
		symbols,
		projectVisibleSymbols,
		activity,
		push,
		udts,
		enums,
		untypedIsVariant: !/^[ \t]*Def(?:Bool|Byte|Int|Lng|LngLng|LngPtr|Cur|Sng|Dbl|Dec|Date|Str|Obj|Var)\b/im.test(source),
		types: moduleTypes(source, mod, activity),
	};
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'VariableGroup') {
			checkDeclarationGroup(ctx, undefined, member);
		}
		if (member.kind === 'Procedure' || member.kind === 'Declare') {
			for (const param of member.params) {
				if (param.paramArray && (param.byVal || param.byRef)) {
					push(
						'paramArrayPassingMode',
						`A ParamArray is always passed ByRef, and takes neither ByVal nor ByRef: '${param.name}'. This is a VBE compile error: Expected: identifier.`,
						param.nameSpan ?? param.span,
					);
				}
			}
		}
		if (member.kind === 'Procedure') {
			const procSym = procedureSymbolFor(symbols, member);
			walkBody(ctx, member, procSym, member.body, []);
		}
	}
}

/** The single variable a bare name binds to, or undefined. */
function variableNamed(ctx: Context, procSym: VbaSymbol | undefined, name: string): VbaSymbol | undefined {
	const binding = sourceIdentifierBinding(ctx.symbols, procSym, ctx.projectVisibleSymbols, name, 'expression');
	if (binding.scope === 'unresolved' || binding.scope === 'ambiguous' || binding.definitions.length !== 1) {
		return undefined;
	}
	const [definition] = binding.definitions;
	return definition.kind === 'localVariable' || definition.kind === 'moduleVariable' || definition.kind === 'parameter'
		? definition
		: undefined;
}

/** The scalar type a Function of the module or project returns, when the name is one. */
function scalarFunctionNamed(ctx: Context, procSym: VbaSymbol | undefined, name: string): string | undefined {
	const binding = sourceIdentifierBinding(ctx.symbols, procSym, ctx.projectVisibleSymbols, name, 'expression');
	if (binding.scope === 'unresolved' || binding.scope === 'ambiguous' || binding.definitions.length !== 1) {
		return undefined;
	}
	const [definition] = binding.definitions;
	const type = normalizeType(definition.asType);
	return (definition.kind === 'function' || (definition.kind === 'declare' && definition.declareKind === 'Function'))
		&& !definition.isArray && type && type !== 'variant' && type !== 'object' && isKnownScalarType(type)
		? definition.asType
		: undefined;
}

function category(ctx: Context, symbol: VbaSymbol | undefined): Category | undefined {
	if (!symbol || symbol.isArray) {
		return undefined;
	}
	const type = normalizeType(symbol.asType);
	if (!type) {
		return ctx.untypedIsVariant ? 'variant' : undefined;
	}
	if (type === 'variant') {
		return 'variant';
	}
	if (NUMBER_TYPES.has(type) || ctx.enums.has(type)) {
		return 'number';
	}
	if (type === 'date' || type === 'string' || type === 'boolean' || type === 'object') {
		return type;
	}
	if (ctx.udts.has(type)) {
		return 'udt';
	}
	// A name no scalar type or Enum claims is a class: Collection, Range, a
	// class module.
	return isKnownScalarType(type) ? undefined : 'class';
}

function walkBody(
	ctx: Context,
	member: ProcedureNode,
	procSym: VbaSymbol | undefined,
	body: readonly BodyNode[],
	counters: readonly string[],
): void {
	for (const node of body) {
		if (isInactiveNode(ctx.activity, node)) {
			continue;
		}
		let inner = counters;
		switch (node.kind) {
			case 'VariableGroup':
				checkDeclarationGroup(ctx, procSym, node);
				break;
			case 'ForBlock':
				inner = checkFor(ctx, procSym, node, counters);
				break;
			case 'SelectBlock':
				checkSelectHead(ctx, node.body);
				break;
			case 'WithBlock':
				checkWithTarget(ctx, procSym, node.span);
				break;
			case 'IfBlock':
				for (const branch of node.branches) {
					checkCondition(ctx, procSym, branch.conditionRaw, branch.headerSpan);
				}
				break;
			case 'Statement':
			case 'Assignment':
			case 'Call':
				checkStatement(ctx, procSym, node.span);
				break;
			default:
				break;
		}
		if ('body' in node && Array.isArray((node as { body?: unknown }).body)) {
			walkBody(ctx, member, procSym, (node as { body: BodyNode[] }).body, inner);
		}
	}
}

function checkFor(
	ctx: Context,
	procSym: VbaSymbol | undefined,
	node: Extract<BodyNode, { kind: 'ForBlock' }>,
	counters: readonly string[],
): readonly string[] {
	const name = node.controlVariable?.trim();
	if (!name || !/^[\p{L}_][\p{L}\p{N}_]*$/u.test(name) || !node.controlVariableSpan) {
		return counters;
	}
	const key = name.toLowerCase();
	if (isConstantName(ctx, procSym, name)) {
		variableRequired(ctx, name, node.controlVariableSpan);
	}
	if (counters.includes(key)) {
		ctx.push(
			'forVariableInUse',
			`'${name}' is already the control variable of an enclosing For. This is a VBE compile error: For control variable already in use.`,
			node.controlVariableSpan,
		);
	}
	const variable = variableNamed(ctx, procSym, name);
	if (!node.each) {
		const kind = category(ctx, variable);
		if (kind === 'string' || kind === 'boolean' || kind === 'object' || kind === 'udt' || kind === 'class') {
			ctx.push(
				'forCounterType',
				`A For counter must be a number, a Date or a Variant, and '${name}' is declared As ${variable!.asType}. This is a VBE compile error: Type mismatch.`,
				node.controlVariableSpan,
			);
		}
	} else if (node.sourceExpression && node.sourceExpressionSpan) {
		const sourceTokens = rawExpressionTokens(node.sourceExpression);
		const sourceName = sourceTokens.length === 1 ? tokenName(sourceTokens[0]) : undefined;
		const array = sourceName ? variableNamed(ctx, procSym, sourceName) : undefined;
		const element = normalizeType(array?.asType);
		if (array?.isArray && (array.fixedLength !== undefined || (element !== undefined && ctx.udts.has(element)))) {
			ctx.push(
				'forEachSourceType',
				`For Each cannot walk '${sourceName}', an array of ${array.fixedLength !== undefined ? 'fixed-length strings' : `the user-defined type ${array.asType}`}. This is a VBE compile error: For Each may not be used on array of user-defined type or fixed-length strings.`,
				node.sourceExpressionSpan,
			);
		}
	}
	return [...counters, key];
}

/** A statement or label between `Select Case` and its first `Case`. */
function checkSelectHead(ctx: Context, body: readonly BodyNode[]): void {
	for (const node of body) {
		if (isInactiveNode(ctx.activity, node) || node.kind === 'ConditionalDirective') {
			continue;
		}
		const toks = statementTokens(ctx.source, node.span).filter((tok) => tok.kind !== 'comment' && tok.kind !== 'newline');
		if (toks.length === 0) {
			continue;
		}
		if (tokenText(toks[0]) === 'case') {
			return;
		}
		ctx.push(
			'statementBeforeFirstCase',
			'Nothing but comments may come between Select Case and its first Case. This is a VBE compile error: Statements and labels invalid between Select Case and first Case.',
			{ start: node.span.start + toks[0].start, end: node.span.start + toks[toks.length - 1].end },
		);
		return;
	}
}

/** The `With` line's tokens, up to the end of that physical line. */
function headerTokens(source: string, span: Span): VbaToken[] {
	const lineEnd = source.slice(span.start, span.end).search(/\r|\n/);
	const toks = statementTokens(source, lineEnd < 0 ? span : { start: span.start, end: span.start + lineEnd });
	const end = toks.findIndex((tok) => tok.kind === 'comment' || tok.kind === 'colon');
	return end < 0 ? toks : toks.slice(0, end);
}

function checkWithTarget(ctx: Context, procSym: VbaSymbol | undefined, span: Span): void {
	const toks = headerTokens(ctx.source, span);
	if (tokenText(toks[0]) !== 'with') {
		return;
	}
	// `With a(1)` on an array of Long, `With TakeL(1)` on a Function that
	// returns one (issue #325, measured in Excel 16.0).
	if (toks.length > 3 && toks[2].rawText === '(' && matchParenFrom(toks, 2) === toks.length - 1 && tokenName(toks[1])) {
		const name = toks[1].rawText;
		const array = variableNamed(ctx, procSym, name);
		const element = array?.isArray ? category(ctx, { ...array, isArray: false }) : undefined;
		const fn = array ? undefined : scalarFunctionNamed(ctx, procSym, name);
		if (element === 'number' || element === 'string' || element === 'boolean' || element === 'date' || fn) {
			ctx.push(
				'withScalarTarget',
				`With needs an object, a user-defined type or a Variant, and '${toks.slice(1).map((tok) => tok.rawText).join('')}' is ${fn ? `a ${fn}, what the Function returns` : 'an element of a scalar array'}. This is a VBE compile error: With object must be user-defined type, Object, or Variant.`,
				{ start: span.start + toks[1].start, end: span.start + toks[toks.length - 1].end },
			);
		}
		return;
	}
	if (toks.length !== 2) {
		return;
	}
	const target = toks[1];
	const literal = target.kind === 'stringLiteral' || target.kind === 'integerLiteral' || target.kind === 'floatLiteral' || target.kind === 'dateLiteral';
	const name = literal ? undefined : tokenName(target);
	const kind = name ? category(ctx, variableNamed(ctx, procSym, name)) : undefined;
	if (literal || kind === 'number' || kind === 'string' || kind === 'boolean' || kind === 'date') {
		ctx.push(
			'withScalarTarget',
			`With needs an object, a user-defined type or a Variant, and '${target.rawText}' is ${literal ? 'a literal' : 'a scalar'}. This is a VBE compile error: With object must be user-defined type, Object, or Variant.`,
			absoluteSpan(span, target),
		);
	}
}

function checkCondition(ctx: Context, procSym: VbaSymbol | undefined, conditionRaw: string | undefined, headerSpan: Span): void {
	const toks = conditionRaw ? rawExpressionTokens(conditionRaw) : [];
	const name = toks.length === 1 ? tokenName(toks[0]) : undefined;
	if (name && category(ctx, variableNamed(ctx, procSym, name)) === 'udt') {
		const at = ctx.source.indexOf(conditionRaw!.trim(), headerSpan.start);
		udtMismatch(ctx, name, at < 0 ? headerSpan : { start: at, end: at + name.length });
	}
}

function udtMismatch(ctx: Context, name: string, span: Span): void {
	ctx.push(
		'udtValueMismatch',
		`'${name}' is a user-defined type, which cannot be used as a single value here. This is a VBE compile error: Type mismatch.`,
		span,
	);
}

function udtCoercion(ctx: Context, name: string, span: Span): void {
	ctx.push(
		'udtVariantCoercion',
		`'${name}' is a user-defined type, and a Variant cannot hold one declared in a standard module or a private class. This is a VBE compile error: Only user-defined types defined in public object modules can be coerced to or from a variant or passed to late-bound functions.`,
		span,
	);
}

function checkStatement(ctx: Context, procSym: VbaSymbol | undefined, span: Span): void {
	const toks = statementTokensAfterLeadingLabel(ctx.source, span).filter((tok) => tok.kind !== 'comment');
	if (toks.length === 0) {
		return;
	}
	checkTypeSuffixes(ctx, procSym, toks, span);
	checkNamedArguments(ctx, procSym, toks, span);
	checkMidTarget(ctx, procSym, toks, span);
	const first = tokenText(toks[0]);
	if (first === 'return' && toks.length > 1 && toks[1].kind !== 'colon') {
		ctx.push(
			'returnWithValue',
			'Return takes no value in VBA: it goes back from a GoSub. Assign the function name instead. This is a VBE compile error: Syntax error.',
			{ start: span.start + toks[0].start, end: span.start + toks[toks.length - 1].end },
		);
		return;
	}
	const udtAt = (i: number): string | undefined => {
		const tok = toks[i];
		const name = tok ? tokenName(tok) : undefined;
		if (!name || toks[i - 1]?.rawText === '.' || toks[i + 1]?.rawText === '.' || toks[i + 1]?.rawText === '(') {
			return undefined;
		}
		return category(ctx, variableNamed(ctx, procSym, name)) === 'udt' ? name : undefined;
	};
	const at = (i: number): Span => absoluteSpan(span, toks[i]);

	if (first === 'set') {
		const name = udtAt(1);
		if (name && toks[2]?.rawText === '=') {
			ctx.push(
				'setRequiresObject',
				`'${name}' is a user-defined type, not an object, so Set cannot assign it. This is a VBE compile error: Object required.`,
				at(1),
			);
		}
		return;
	}
	if (first === 'debug' && toks[1]?.rawText === '.' && tokenText(toks[2]) === 'print') {
		for (let i = 3; i < toks.length; i++) {
			const name = udtAt(i);
			if (name && isArgumentBoundary(toks[i - 1]) && isArgumentBoundary(toks[i + 1])) {
				udtMismatch(ctx, name, at(i));
			}
		}
	}
	// `LSet n = 5` on a Long, `RSet a = b` on a Type (issue #451, measured
	// in Excel 16.0). LSet takes a String or a Type, and runs on a Variant;
	// RSet takes a String or a Variant only.
	if ((first === 'lset' || first === 'rset') && toks[2]?.rawText === '=') {
		const target = variableNamed(ctx, procSym, tokenName(toks[1]) ?? '');
		const kind = category(ctx, target);
		const refused = kind === 'number' || kind === 'boolean' || kind === 'date' || (first === 'rset' && kind === 'udt');
		if (target && refused && !target.isArray) {
			ctx.push(
				'lsetTypeMismatch',
				`'${toks[1].rawText}' is ${/^[aeiou]/i.test(target.asType ?? '') ? 'an' : 'a'} ${target.asType}, which ${toks[0].rawText} cannot fill. This is a VBE compile error: ${first === 'lset' ? 'LSet allowed only on strings and user-defined types' : 'RSet allowed only on strings'}.`,
				{ start: span.start + toks[0].start, end: span.start + toks[1].end },
			);
			return;
		}
	}
	if (first === 'lset' && toks.length === 4 && toks[2].rawText === '=') {
		checkLSet(ctx, procSym, toks, span);
		return;
	}
	checkMethodArguments(ctx, procSym, toks, span);
	if (first === 'msgbox') {
		const start = toks[1]?.rawText === '(' ? 2 : 1;
		const name = udtAt(start);
		if (name && isArgumentBoundary(toks[start + 1])) {
			udtCoercion(ctx, name, at(start));
		}
		return;
	}

	// `a = b`: the target and the value's kinds, when each is a bare name or a
	// literal.
	const eq = first === 'let' ? 2 : 1;
	if (toks[eq]?.rawText === '=' && toks.length === eq + 2 && tokenName(toks[eq - 1])) {
		const targetName = tokenName(toks[eq - 1])!;
		const target = variableNamed(ctx, procSym, targetName);
		const targetKind = category(ctx, target);
		const value = toks[eq + 1];
		const valueName = tokenName(value);
		const valueSymbol = valueName ? variableNamed(ctx, procSym, valueName) : undefined;
		const valueKind = category(ctx, valueSymbol);
		const literal = value.kind === 'stringLiteral' || value.kind === 'integerLiteral' || value.kind === 'floatLiteral' || value.kind === 'dateLiteral';
		if (targetKind === 'udt') {
			const otherType = valueKind === 'udt' && normalizeType(valueSymbol!.asType) !== normalizeType(target!.asType);
			if (literal || otherType || valueKind === 'number' || valueKind === 'string' || valueKind === 'boolean' || valueKind === 'date') {
				udtMismatch(ctx, targetName, at(eq - 1));
			}
		} else if (valueKind === 'udt') {
			if (targetKind === 'variant') {
				udtCoercion(ctx, valueName!, at(eq + 1));
			} else if (targetKind === 'number' || targetKind === 'string' || targetKind === 'boolean' || targetKind === 'date') {
				udtMismatch(ctx, valueName!, at(eq + 1));
			}
		}
		return;
	}

	// A user-defined type beside an operator is used as a value: `t + 1`,
	// `t = u` in an expression. The assignment's own `=` is not one, however
	// long its target: `parents(depth) = cInfo`, `.Items(0).Id = rec`.
	const assignmentEq = startsAssignment(toks) ? firstTopLevelEquals(toks) : -1;
	for (let i = 0; i < toks.length; i++) {
		const name = udtAt(i);
		if (!name) {
			continue;
		}
		const before = toks[i - 1];
		const after = toks[i + 1];
		const operatorBefore = before && i - 1 !== assignmentEq && isOperator(before);
		const operatorAfter = after && i + 1 !== assignmentEq && isOperator(after);
		if (operatorBefore || operatorAfter) {
			udtMismatch(ctx, name, at(i));
		}
	}
}

/**
 * Whether a statement is an assignment by its first token: a name, a
 * `.member`, Let, Set, or LSet and RSet, which copy one user-defined type
 * into another.
 */
function startsAssignment(toks: readonly VbaToken[]): boolean {
	const first = toks[0];
	if (!first) {
		return false;
	}
	const word = tokenText(first);
	return first.rawText === '.' || ['let', 'set', 'lset', 'rset'].includes(word)
		|| first.kind === 'identifier' || first.kind === 'bracketedIdentifier';
}

function firstTopLevelEquals(toks: readonly VbaToken[]): number {
	let depth = 0;
	for (let i = 0; i < toks.length; i++) {
		const raw = toks[i].rawText;
		depth += raw === '(' ? 1 : raw === ')' ? -1 : 0;
		if (depth === 0 && raw === '=') {
			return i;
		}
	}
	return -1;
}

function isOperator(tok: VbaToken): boolean {
	// `Item:=t` names an argument, which udt-variant-coercion reads.
	return tok.kind === 'operator' ? tok.rawText !== '.' && tok.rawText !== '!' && tok.rawText !== ':=' : OPERATOR_WORDS.has(tokenText(tok));
}

/**
 * `LSet a = b` between two Types (issue #253): refused when they differ and
 * either holds a member that is not of fixed size.
 */
function checkLSet(ctx: Context, procSym: VbaSymbol | undefined, toks: readonly VbaToken[], span: Span): void {
	const target = variableNamed(ctx, procSym, tokenName(toks[1]) ?? '');
	const value = variableNamed(ctx, procSym, tokenName(toks[3]) ?? '');
	const targetType = normalizeType(target?.asType);
	const valueType = normalizeType(value?.asType);
	if (category(ctx, target) !== 'udt' || category(ctx, value) !== 'udt' || !targetType || !valueType || targetType === valueType) {
		return;
	}
	const held = variableSizeMember(ctx, targetType, 0) ?? variableSizeMember(ctx, valueType, 0);
	if (!held) {
		return;
	}
	ctx.push(
		'lsetTypeMismatch',
		`LSet copies '${toks[3].rawText}' (${value!.asType}) into '${toks[1].rawText}' (${target!.asType}), two different user-defined types, and the member ${held}. LSet copies between two types only when every member is of fixed size. This is a VBE compile error: Type mismatch.`,
		{ start: span.start + toks[0].start, end: span.start + toks[3].end },
	);
}

/** A member of a module Type that is not of fixed size, described; undefined when every member is, or one is unknown. */
function variableSizeMember(ctx: Context, type: string, depth: number): string | undefined {
	const fields = ctx.types.get(type);
	if (!fields || depth > 8) {
		return undefined;
	}
	for (const field of fields.values()) {
		const fieldType = normalizeType(field.type);
		if (field.isArray && !isFixedArrayField(field)) {
			return `'${field.name}' is a dynamic array`;
		}
		if (!fieldType || fieldType === 'variant') {
			return `'${field.name}' is a Variant`;
		}
		if (fieldType === 'string' && field.fixedLength === undefined) {
			return `'${field.name}' is a variable-length String`;
		}
		if (ctx.types.has(fieldType)) {
			const nested = variableSizeMember(ctx, fieldType, depth + 1);
			if (nested) {
				return nested;
			}
		} else if (fieldType === 'object' || fieldType === 'collection') {
			return `'${field.name}' is an object`;
		}
	}
	return undefined;
}

/** VBA's functions measured to refuse a user-defined type for their Variant parameter (issue #253). */
const VARIANT_FUNCTIONS: ReadonlySet<string> = new Set([
	'array', 'choose', 'cvar', 'format', 'iif', 'isarray', 'isdate', 'isempty', 'iserror', 'ismissing',
	'isnull', 'isnumeric', 'isobject', 'typename', 'vartype',
]);

/**
 * `VBA.TypeName(t)`: the qualified forms of VBA's functions refuse a Type as
 * the bare ones do; argument-shape-mismatch reads the bare ones.
 */
function checkQualifiedLibraryArguments(ctx: Context, procSym: VbaSymbol | undefined, toks: readonly VbaToken[], span: Span): void {
	for (let i = 0; i + 3 < toks.length; i++) {
		if (tokenText(toks[i]) !== 'vba' || toks[i - 1]?.rawText === '.' || toks[i + 1].rawText !== '.' || !VARIANT_FUNCTIONS.has(tokenText(toks[i + 2])) || toks[i + 3].rawText !== '(') {
			continue;
		}
		let depth = 0;
		for (let j = i + 4; j < toks.length; j++) {
			const raw = toks[j].rawText;
			if (raw === '(') {
				depth++;
			} else if (raw === ')') {
				if (depth-- === 0) {
					break;
				}
			} else if (depth === 0 && tokenName(toks[j]) && [',', '('].includes(toks[j - 1].rawText) && [',', ')'].includes(toks[j + 1]?.rawText ?? '')
				&& category(ctx, variableNamed(ctx, procSym, tokenName(toks[j])!)) === 'udt') {
				udtCoercion(ctx, toks[j].rawText, absoluteSpan(span, toks[j]));
			}
		}
	}
}

/**
 * `c.Add t`, `o.Add Item:=t`, `Call c.Add(t)`: a user-defined type passed to
 * a method of a Collection, an Object or a Variant (issue #253).
 */
function checkMethodArguments(ctx: Context, procSym: VbaSymbol | undefined, toks: readonly VbaToken[], span: Span): void {
	checkQualifiedLibraryArguments(ctx, procSym, toks, span);
	const at = tokenText(toks[0]) === 'call' ? 1 : 0;
	const receiver = tokenName(toks[at]) ? variableNamed(ctx, procSym, tokenName(toks[at])!) : undefined;
	const receiverType = normalizeType(receiver?.asType) ?? (ctx.untypedIsVariant ? 'variant' : undefined);
	if (!receiver || receiver.isArray || !['collection', 'object', 'variant'].includes(receiverType ?? '') || toks[at + 1]?.rawText !== '.' || !tokenName(toks[at + 2])) {
		return;
	}
	let start = at + 3;
	let end = toks.length;
	if (toks[start]?.rawText === '(' && toks[toks.length - 1]?.rawText === ')') {
		start++;
		end--;
	}
	let depth = 0;
	let slotStart = start;
	for (let i = start; i <= end; i++) {
		const raw = toks[i]?.rawText;
		if (i < end && raw !== ',') {
			depth += raw === '(' ? 1 : raw === ')' ? -1 : 0;
			continue;
		}
		if (depth > 0) {
			continue;
		}
		const slot = toks.slice(slotStart, i);
		slotStart = i + 1;
		const value = slot[1]?.rawText === ':=' ? slot.slice(2) : slot;
		const name = value.length === 1 ? tokenName(value[0]) : undefined;
		if (name && category(ctx, variableNamed(ctx, procSym, name)) === 'udt') {
			udtCoercion(ctx, name, absoluteSpan(span, value[0]));
		}
	}
}

/** Whether a token ends or begins a Print or MsgBox argument. */
function isArgumentBoundary(tok: VbaToken | undefined): boolean {
	return !tok || tok.rawText === ',' || tok.rawText === ';' || tok.rawText === ')' || tok.rawText === '(' || tok.kind === 'colon'
		|| tokenText(tok) === 'print';
}

/** A name that binds to a Const or an Enum member. */
function isConstantName(ctx: Context, procSym: VbaSymbol | undefined, name: string): boolean {
	const binding = sourceIdentifierBinding(ctx.symbols, procSym, ctx.projectVisibleSymbols, name, 'expression');
	return binding.scope !== 'ambiguous' && binding.scope !== 'unresolved'
		&& binding.definitions.length > 0
		&& binding.definitions.every((definition) => definition.kind === 'constant' || definition.kind === 'enumMember');
}

function variableRequired(ctx: Context, name: string, span: Span): void {
	ctx.push(
		'variableRequired',
		`'${name}' is a constant, and a constant cannot be assigned to here. This is a VBE compile error: Variable required - can't assign to this expression.`,
		span,
	);
}

/**
 * A Const whose value, or a Dim whose bounds, name a variable: "Constant
 * expression required". A ReDim is a statement and never reaches here.
 */
function checkDeclarationGroup(
	ctx: Context,
	procSym: VbaSymbol | undefined,
	group: Extract<BodyNode, { kind: 'VariableGroup' }>,
): void {
	for (const decl of group.declarations) {
		if (isInactiveNode(ctx.activity, decl)) {
			continue;
		}
		const toks = statementTokens(ctx.source, decl.span);
		let read: VbaToken[] = [];
		if (group.isConst) {
			const eq = toks.findIndex((tok) => tok.rawText === '=');
			read = eq < 0 ? [] : toks.slice(eq + 1);
		} else if (decl.fixedLength !== undefined) {
			// `Dim s As String * L` (issue #451, measured in Excel 16.0).
			const star = toks.findIndex((tok) => tok.rawText === '*');
			read = star < 0 ? [] : toks.slice(star + 1);
		} else if (decl.isArray && (decl.arrayBounds ?? '').trim() !== '') {
			const open = toks.findIndex((tok) => tok.rawText === '(');
			let depth = 0;
			for (let i = open; open >= 0 && i < toks.length; i++) {
				depth += toks[i].rawText === '(' ? 1 : toks[i].rawText === ')' ? -1 : 0;
				if (depth === 0) {
					read = toks.slice(open + 1, i);
					break;
				}
			}
		}
		for (let i = 0; i < read.length; i++) {
			const name = tokenName(read[i]);
			if (!name || read[i - 1]?.rawText === '.' || read[i + 1]?.rawText === '.' || read[i + 1]?.rawText === '(') {
				continue;
			}
			if (!variableNamed(ctx, procSym, name)) {
				continue;
			}
			ctx.push(
				group.isConst ? 'constValueNotConstant' : 'arrayBoundNotConstant',
				group.isConst
					? `Const '${decl.name}' takes its value from the variable '${name}'. This is a VBE compile error: Constant expression required.`
					: decl.fixedLength !== undefined
						? `The length of the fixed-length String '${decl.name}' names the variable '${name}'; it must be a constant. This is a VBE compile error: Constant expression required.`
						: `The bounds of '${decl.name}' name the variable '${name}'; a Dim needs constants there, and ReDim takes a variable. This is a VBE compile error: Constant expression required.`,
				absoluteSpan(decl.span, read[i]),
			);
		}
	}
}

/** The type each type-declaration character stands for. */
const SUFFIX_TYPES: Readonly<Record<string, string>> = {
	'%': 'integer', '&': 'long', '^': 'longlong', '@': 'currency', '!': 'single', '#': 'double', '$': 'string',
};

/** `n% = 2` with n As Long: the character names another type than the declaration. */
function checkTypeSuffixes(ctx: Context, procSym: VbaSymbol | undefined, toks: readonly VbaToken[], span: Span): void {
	for (let i = 0; i + 1 < toks.length; i++) {
		const nameTok = toks[i];
		const suffixTok = toks[i + 1];
		const suffixType = SUFFIX_TYPES[suffixTok.rawText];
		const name = tokenName(nameTok);
		if (!suffixType || !name || suffixTok.start !== nameTok.end || toks[i - 1]?.rawText === '.') {
			continue;
		}
		// `rs!Field` is a member, not a suffix.
		const after = toks[i + 2];
		if (after && after.start === suffixTok.end && (tokenName(after) || after.kind === 'integerLiteral')) {
			continue;
		}
		const variable = variableNamed(ctx, procSym, name);
		if (!variable || variable.isArray) {
			continue;
		}
		const declared = normalizeType(variable.asType) ?? (ctx.untypedIsVariant ? 'variant' : undefined);
		if (!declared || declared === suffixType) {
			continue;
		}
		ctx.push(
			'typeSuffixMismatch',
			`'${name}${suffixTok.rawText}' says ${suffixType}, but '${name}' is declared ${variable.asType ?? 'Variant'}. This is a VBE compile error: Type-declaration character does not match declared data type.`,
			{ start: span.start + nameTok.start, end: span.start + suffixTok.end },
		);
	}
}

/** The functions that take no named arguments: the ones VBA compiles as keywords. */
const NO_NAMED_ARGUMENTS = new Set([
	'instr', 'instrb', 'len', 'lenb', 'strcomp', 'abs', 'int', 'fix', 'sgn',
	'cstr', 'cint', 'clng', 'cdbl', 'cbool', 'cdate', 'cvar', 'cbyte', 'ccur', 'csng', 'clnglng', 'clngptr',
]);

function checkNamedArguments(ctx: Context, procSym: VbaSymbol | undefined, toks: readonly VbaToken[], span: Span): void {
	for (let i = 0; i + 1 < toks.length; i++) {
		const name = tokenName(toks[i]);
		if (!name || !NO_NAMED_ARGUMENTS.has(name.toLowerCase()) || toks[i + 1].rawText !== '(' || toks[i - 1]?.rawText === '.') {
			continue;
		}
		// A project procedure of the same name takes named arguments like any other.
		const binding = sourceIdentifierBinding(ctx.symbols, procSym, ctx.projectVisibleSymbols, name, 'call');
		if (binding.scope !== 'unresolved') {
			continue;
		}
		let depth = 0;
		for (let j = i + 1; j < toks.length; j++) {
			depth += toks[j].rawText === '(' ? 1 : toks[j].rawText === ')' ? -1 : 0;
			if (depth === 0) {
				break;
			}
			if (depth === 1 && toks[j].rawText === ':=') {
				ctx.push(
					'namedArgumentNotAllowed',
					`${name} takes its arguments by position only. This is a VBE compile error: Syntax error.`,
					{ start: span.start + toks[j - 1].start, end: span.start + toks[j].end },
				);
				break;
			}
		}
	}
}

/** `Mid$(S, 1, 1) = "x"` with S a Const. */
function checkMidTarget(ctx: Context, procSym: VbaSymbol | undefined, toks: readonly VbaToken[], span: Span): void {
	const head = tokenText(toks[0]);
	if (head !== 'mid' && head !== 'midb') {
		return;
	}
	const open = toks[1]?.rawText === '$' ? 2 : 1;
	const target = toks[open + 1];
	const name = target ? tokenName(target) : undefined;
	if (toks[open]?.rawText !== '(' || !name || toks[open + 2]?.rawText !== ',') {
		return;
	}
	let depth = 0;
	for (let j = open; j < toks.length; j++) {
		depth += toks[j].rawText === '(' ? 1 : toks[j].rawText === ')' ? -1 : 0;
		if (depth === 0) {
			if (toks[j + 1]?.rawText === '=' && isConstantName(ctx, procSym, name)) {
				variableRequired(ctx, name, absoluteSpan(span, target!));
			}
			return;
		}
	}
}
