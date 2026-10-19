// `TypeOf <expr> Is <Type>` always-False detection (MS-VBAL 5.6 TypeOf...Is).
//
// `TypeOf x Is Y` tests whether x's runtime object is-a Y. When x is declared as
// a known, concrete object type A that can never hold a Y, the test is provably
// always False - a dead branch and almost always a bug.
//
// This is the first type rule to consume the §5.6 expression AST directly: it
// reads the parsed `TypeOfIsExpr` nodes that Slices 1-3 produce (notably in
// `If TypeOf x Is Y Then` branch conditions, now parsed expressions) instead of
// re-tokenizing. It reuses the existing object-compatibility tables
// (`objectAssignmentIncompatibilityReason`, which already encodes host/project
// identity and `Implements` interface relationships) so no new type matrix is
// introduced.
//
// No-false-positive discipline:
//   - Only IdentifierExpr operands with a known declared OBJECT type are checked.
//   - `Object`/`Variant`/unknown/scalar operands or targets stay quiet.
//   - The operand type A must be CONCRETE: a host type, or a project class that
//     no project class `Implements`. An interface-typed operand could hold a
//     subtype that is-a Y, so it stays quiet (avoids the diamond false positive).
//   - Fires only when A and Y are mutually assignment-incompatible (neither is-a
//     the other), proven via the existing object tables in both directions.

import type {
	ExprNode,
	IdentifierExpr,
	Span,
	ProcedureNode,
	TypeOfIsExpr,
} from '../../parser/nodes';
import type { MemberCompletionContext } from '../../completion/memberAccess';
import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { VbaSymbol } from '../../symbols/symbolModel';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import { procedureSymbolFor, type PushFn } from '../analysisContext';
import {
	createObjectAssignmentTypeResolver,
	createObjectTypeImplementationLookup,
	createProjectInterfaceSharingLookup,
	isKnownScalarType,
	isNumericType,
	constantStringValue,
	sourceIdentifierBinding,
	sourceNameScopeFor,
	type DeclaredValueShape,
	objectAssignmentIncompatibilityReason,
	typeEnvironmentFor,
} from '../typeInference';
import { tokenizeCached } from '../../lexer/tokenize';
import { expressionShapeResolver, type ExpressionShapeContext } from '../expressionShape';
import type { ProcedureExpressionVisitor } from '../exprWalk';
import { retainedExpressionText } from './binaryOperandScalar';

/**
 * Rule: `TypeOf` requires an object expression before `Is` (`TypeOf x Is Y`). A
 * malformed `TypeOf Is Y` (no operand) is a compile error - VBE "Syntax error"
 * (oracle-verified `corpus_excel_syntax_006_compile`). Detected by a whole-source
 * scan for an adjacent `TypeOf` `Is` pair (an operand always sits between them in
 * valid code); inactive conditional-compilation regions are skipped (no-FP).
 */
export function checkTypeOfMissingOperand(
	source: string,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const toks = tokenizeCached(source).filter((t) => t.kind !== 'comment' && t.kind !== 'newline');
	for (let i = 0; i + 1 < toks.length; i++) {
		const word = (toks[i].canonicalText ?? toks[i].rawText).toLowerCase();
		if (word !== 'typeof') {
			continue;
		}
		const nextWord = (toks[i + 1].canonicalText ?? toks[i + 1].rawText).toLowerCase();
		if (nextWord !== 'is') {
			continue;
		}
		const span = { start: toks[i].start, end: toks[i + 1].end };
		if (activity?.isInactive(span)) {
			continue;
		}
		push("typeofMissingOperand", "'TypeOf' requires an object expression before 'Is'.", span);
	}
}

interface TypeOfQueries {
	resolveType: ReturnType<typeof createObjectAssignmentTypeResolver>;
	shareInterfaces: ReturnType<typeof createProjectInterfaceSharingLookup>;
	implementsType: ReturnType<typeof createObjectTypeImplementationLookup>;
	isImplemented: (key: string, display: string) => boolean;
}

export function checkTypeOfIsCompatibility(
	symbols: ReturnType<typeof buildModuleSymbols>,
	memberCtx: MemberCompletionContext,
	push: PushFn,
): ProcedureExpressionVisitor {
	// Project metadata is stable for this analysis query. Keep all derived
	// lookups lazy and query-owned so the next query observes fresh metadata.
	let implementedNames: Set<string> | undefined;
	const queries: TypeOfQueries = {
		resolveType: createObjectAssignmentTypeResolver(memberCtx),
		shareInterfaces: createProjectInterfaceSharingLookup(memberCtx),
		implementsType: createObjectTypeImplementationLookup(),
		isImplemented: (key, display) => {
			if (!implementedNames) {
				implementedNames = new Set();
				// The original exclusion considered every supplied surface kind.
				for (const type of memberCtx.projectClassMembers ?? []) {
					for (const name of type.implements ?? []) { implementedNames.add(name.toLowerCase()); }
				}
			}
			return implementedNames.has(key) || implementedNames.has(display.toLowerCase());
		},
	};
	return (member) => {
		const env = typeEnvironmentFor(symbols, member);
		return (expr) => {
			if (expr.exprKind === 'TypeOfIsExpr') {
				checkTypeOfIs(expr, env, memberCtx, queries, push);
			}
		};
	};
}

function checkTypeOfIs(
	expr: TypeOfIsExpr,
	env: ReadonlyMap<string, string>,
	memberCtx: MemberCompletionContext,
	queries: TypeOfQueries,
	push: PushFn,
): void {
	if (expr.operand.exprKind !== 'IdentifierExpr') {
		return; // v1: only simple identifier operands have a known declared type
	}
	const operandName = (expr.operand as IdentifierExpr).name;
	const declared = env.get(operandName.toLowerCase());
	if (!declared) {
		return; // undeclared / unknown type -> quiet
	}
	const operandType = queries.resolveType(declared);
	const targetType = queries.resolveType(expr.typeName);
	if (!operandType || !targetType) {
		return; // not both known object types -> quiet
	}
	if (operandType.kind === 'generic' || targetType.kind === 'generic') {
		return; // Object operand or `Is Object` (always True) -> quiet
	}
	if (operandType.key === targetType.key) {
		return; // same type (always True) -> quiet
	}
	// Concrete-operand gate: an interface-typed operand could hold a subtype that
	// is-a the target, so only fire when A cannot be an interface implemented by
	// some class (host types are never user-implementable).
	if (operandType.kind === 'project' && queries.isImplemented(operandType.key, operandType.display)) {
		return;
	}
	// Mutual incompatibility: neither type is assignable to the other (reuses the
	// existing object tables, which already honour `Implements`). If either is
	// assignable, `TypeOf` could be True, so stay quiet.
	const operandCanBeTarget = objectAssignmentIncompatibilityReason(
		expr.typeName,
		{ type: declared, label: declared, span: expr.span },
		memberCtx,
		queries.resolveType,
		queries.shareInterfaces,
		queries.implementsType,
	) === undefined;
	const targetCanBeOperand = objectAssignmentIncompatibilityReason(
		declared,
		{ type: expr.typeName, label: expr.typeName, span: expr.span },
		memberCtx,
		queries.resolveType,
		queries.shareInterfaces,
		queries.implementsType,
	) === undefined;
	if (operandCanBeTarget || targetCanBeOperand) {
		return;
	}
	push(
		'typeOfIsAlwaysFalse',
		`'TypeOf ... Is ${targetType.display}' is always False: '${operandName}' is declared As ${operandType.display}, which is never ${targetType.display}.`,
		expr.span,
	);
}

/**
 * Rule: the binary `Is` object-identity operator requires object-reference
 * operands on both sides (MS-VBAL 5.6). When an operand is PROVABLY a non-object
 * - a scalar value literal, or an identifier declared As a known scalar type
 * (numeric / String / Boolean / Date) - `a Is b` is a VBE compile error
 * ("Type mismatch", including `cl_ctx_is_long__debug_print_compile`).
 * Oracle-verified rejected at compile:
 * `is_scalar_long_var_compile`, `is_scalar_string_var_compile`,
 * `is_two_scalar_vars_compile`, `is_literal_integer_operands_probe`,
 * `is_literal_string_operands_probe`; with the Variant control
 * `is_variant_var_ok_compile` and the Object control `is_object_var_vs_nothing_probe`
 * accepted.
 *
 * No-false-positive discipline: fires only when an operand is PROVABLY scalar.
 * Variant (can hold an object reference), Object/class-typed operands,
 * `Nothing`/`Null`/`Empty` literals and unknown values stay quiet. Parentheses,
 * declared array elements, UDT fields and known scalar call results are resolved. `TypeOf x Is Y` is a
 * distinct `TypeOfIsExpr`, not a binary `Is`, so it is unaffected.
 */
export function checkIsOperatorOperands(
	symbols: ReturnType<typeof buildModuleSymbols>,
	push: PushFn,
	projectVisibleSymbols?: readonly VbaSymbol[],
	shapeContext?: ExpressionShapeContext,
): ProcedureExpressionVisitor {
	const forProcedure = expressionShapeResolver(symbols, projectVisibleSymbols, shapeContext);
	const isEnum = enumTypeResolver(symbols, projectVisibleSymbols, shapeContext);
	const intrinsicFor = measuredIntrinsicResolver(symbols, projectVisibleSymbols, shapeContext);
	return (member) => {
		const declaredShape = forProcedure(member);
		const intrinsic = intrinsicFor(member);
		const procSym = procedureSymbolFor(symbols, member);
		const values = new Map<string, DeclaredValueShape | undefined>();
		const shapeOf = (expr: ExprNode): DeclaredValueShape | undefined => {
			const shape = declaredShape(expr);
			if (expr.exprKind !== 'IdentifierExpr' || shape?.asType || shape?.isArray || shape?.isScalar) { return shape; }
			const name = expr.name.toLowerCase();
			if (values.has(name)) { return values.get(name) ?? shape; }
			// Use the indexed source binding so local Object/Variant shadows win.
			const binding = sourceIdentifierBinding(symbols, procSym, projectVisibleSymbols, name, 'expression');
			const symbol = binding.definitions.length === 1 ? binding.definitions[0] : undefined;
			const asType = symbol?.kind === 'enumMember' ? symbol.containerName
				: symbol && constantStringValue(symbol) !== undefined ? 'String' : undefined;
			const value = asType ? { asType, isArray: false, isFixedArray: false } : undefined;
			values.set(name, value);
			return value ?? shape;
		};
		return (expr) => {
			if (expr.exprKind !== 'BinaryExpr' || expr.operator !== 'Is') {
				return;
			}
			const offender = nonObjectOperand(expr.left, shapeOf, isEnum, intrinsic, member) ?? nonObjectOperand(expr.right, shapeOf, isEnum, intrinsic, member);
			if (offender) {
				push(
					'isOperatorNonObject',
					`The 'Is' operator requires object operands, but ${offender.detail}, which is not an object. This is a VBE compile error: Type mismatch.`,
					offender.span,
				);
			}
		};
	};
}

/** Describes a provably non-object (scalar) operand of `Is`, or undefined. */
function nonObjectOperand(
	expr: ExprNode,
	shapeOf: ReturnType<ReturnType<typeof expressionShapeResolver>>,
	isEnum: (name: string) => boolean,
	intrinsic: (expr: ExprNode) => string | undefined,
	member: ProcedureNode,
): { span: Span; detail: string } | undefined {
	if (expr.exprKind === 'LiteralExpr') {
		const k = expr.literalKind;
		if (k === 'integer' || k === 'float' || k === 'string' || k === 'date' || k === 'boolean') {
			return { span: expr.span, detail: `'${expr.raw}' is a ${k} literal` };
		}
		return undefined; // Nothing / Null / Empty -> not provably scalar
	}
	if (expr.exprKind === 'ParenExpr') { return nonObjectOperand(expr.inner, shapeOf, isEnum, intrinsic, member); }
	// Only the measured arithmetic forms, with provably scalar inputs. An
	// unknown/Variant operand remains unknown, including inside parentheses.
	if ((expr.exprKind === 'UnaryExpr' || expr.exprKind === 'BinaryExpr') && numericIsOperand(expr, shapeOf)) {
		return { span: expr.span, detail: 'the expression has a scalar numeric type' };
	}
	const shape = shapeOf(expr);
	const declared = shape?.asType;
	if (shape?.isArray || shape?.isScalar || declared && (isKnownScalarType(declared.trim().toLowerCase()) || isEnum(declared.toLowerCase()))) {
		const name = expr.exprKind === 'IdentifierExpr' ? `'${expr.name}${expr.typeSuffix ?? ''}'` : 'the expression';
		return { span: measuredIsSpan(expr, intrinsic(expr), shapeOf, member), detail: `${name} has ${shape?.isArray ? 'array type' : shape?.isScalar ? 'a scalar numeric type' : `type ${declared}`}` };
	}
	return undefined;
}

function numericIsOperand(expr: ExprNode, shapeOf: ReturnType<ReturnType<typeof expressionShapeResolver>>): boolean {
	if (expr.exprKind === 'ParenExpr') { return numericIsOperand(expr.inner, shapeOf); }
	if (expr.exprKind === 'LiteralExpr') { return expr.literalKind === 'integer' || expr.literalKind === 'float'; }
	if (expr.exprKind === 'UnaryExpr' && expr.operator === '-') { return numericIsOperand(expr.operand, shapeOf); }
	if (expr.exprKind === 'BinaryExpr' && expr.operator === '+') { return numericIsOperand(expr.left, shapeOf) && numericIsOperand(expr.right, shapeOf); }
	const shape = shapeOf(expr);
	return !shape?.isArray && shape?.asType !== undefined && isNumericType(shape.asType.toLowerCase());
}

/** Excel selects the callable name (including its member dot), not its arguments. */
function compileOperandSpan(expr: ExprNode, member: ProcedureNode): Span {
	if (expr.exprKind === 'IndexExpr') { return compileOperandSpan(expr.callee, member); }
	if (expr.exprKind === 'MemberAccessExpr' && expr.object && expr.accessKind !== 'bang') {
		const gap = { start: expr.object.span.end, end: expr.memberSpan.start };
		const text = retainedExpressionText(member.body, gap);
		const dot = text?.match(/^\s*\./);
		// Adjacent dots need no source lookup; otherwise prefer the exact name
		// when this AST node does not retain the intervening source text.
		const start = dot ? gap.start + dot[0].length - 1
			: gap.end - gap.start === 1 ? gap.start : expr.memberSpan.start;
		return { start, end: expr.memberSpan.end };
	}
	return expr.span;
}

/** Compile constraints of TypeOf, separate from the compatibility warning. */
export function checkTypeOfExpression(symbols: ReturnType<typeof buildModuleSymbols>, push: PushFn, projectVisibleSymbols?: readonly VbaSymbol[], shapeContext?: ExpressionShapeContext): ProcedureExpressionVisitor {
	const forProcedure = expressionShapeResolver(symbols, projectVisibleSymbols, shapeContext);
	const isEnum = enumTypeResolver(symbols, projectVisibleSymbols, shapeContext);
	const intrinsicFor = measuredIntrinsicResolver(symbols, projectVisibleSymbols, shapeContext);
	return (member) => {
		const shapeOf = forProcedure(member);
		const intrinsic = intrinsicFor(member);
		return (expr) => {
			if (expr.exprKind !== 'TypeOfIsExpr') { return; }
			const shape = shapeOf(expr.operand);
			const builtin = intrinsic(expr.operand);
			// These measured comparison calls compile under TypeOf, despite
			// their scalar constraints in binary Is. No return type is inferred.
			const acceptedComparison = measuredComparisonCall(expr.operand, builtin);
			const operandKind = expr.operand.exprKind;
			const invalidSyntax = builtin && shape && TYPEOF_SYNTAX_INTRINSICS.has(builtin)
				|| operandKind === 'ParenExpr' || operandKind === 'LiteralExpr'
				|| operandKind === 'UnaryExpr' || operandKind === 'BinaryExpr'
				|| isKnownScalarType(expr.typeName.toLowerCase()) || expr.typeName.toLowerCase() === 'variant';
			if (invalidSyntax) {
				push('typeOfInvalidExpression', "TypeOf requires an unparenthesised object expression and an object type after Is. This is a VBE compile error: Syntax error.", expr.span);
			} else if (!acceptedComparison && (shape?.isArray || shape?.isScalar || shape?.asType && (isKnownScalarType(shape.asType.toLowerCase()) || isEnum(shape.asType.toLowerCase())))) {
				push('typeOfInvalidExpression', "TypeOf requires an object expression. This is a VBE compile error: Type mismatch.", compileOperandSpan(expr.operand, member));
			}
		};
	};
}

/** Resolve valid Enum identity once per query, retaining ambiguous source
 * bindings. A matching enum spelling alone is insufficient. */
function enumTypeResolver(
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	context: ExpressionShapeContext | undefined,
): (name: string) => boolean {
	const cache = new Map<string, boolean>();
	let classNames: Set<string> | undefined;
	return (name) => {
		const key = name.trim().toLowerCase();
		const cached = cache.get(key);
		if (cached !== undefined) { return cached; }
		// Excel rejects module-qualified source Enum declarations (measured
		// imported_enum_35/36). The declaration diagnostic owns that error;
		// do not derive a secondary operand constraint from an invalid type.
		if (key.includes('.')) { cache.set(key, false); return false; }
		const definitions = sourceIdentifierBinding(symbols, undefined, projectVisibleSymbols, key, 'typeName').definitions;
		classNames ??= new Set((context?.projectClassMembers ?? []).filter((type) => ['class', 'document', 'userform'].includes(type.kind)).map((type) => type.name.toLowerCase()));
		const result = definitions.length === 1 && definitions[0].kind === 'enum'
			&& (definitions[0].moduleName.toLowerCase() === symbols.moduleName.toLowerCase() || !classNames.has(key));
		cache.set(key, result);
		return result;
	};
}

// Unqualified grammar observations from gap_contexts_review_20261019 and
// gap_contexts_validation_20261019. Qualification is intentionally excluded:
// VBA's qualified call grammar need not share these intrinsic productions.
const TYPEOF_SYNTAX_INTRINSICS = new Set([
	'string$', 'cdbl', 'doevents', 'lenb', 'lbound', 'seek', 'cbool',
	'cbyte', 'ccur', 'clnglng', 'csng', 'cvar', 'array', 'cverr',
	'abs', 'int', 'fix', 'len', 'cdate', 'cdec', 'cint', 'clng', 'cstr', 'sgn', 'ubound',
]);

function measuredIntrinsicResolver(
	symbols: ReturnType<typeof buildModuleSymbols>,
	visible: readonly VbaSymbol[] | undefined,
	context: ExpressionShapeContext | undefined,
): (member: ProcedureNode) => (expr: ExprNode) => string | undefined {
	let externalNames: Set<string> | undefined;
	return member => {
		let scope: ReturnType<typeof sourceNameScopeFor> | undefined;
		return expr => {
			const callee = expr.exprKind === 'IndexExpr' ? expr.callee : expr;
			if (callee.exprKind !== 'IdentifierExpr') { return undefined; }
			const name = callee.name.toLowerCase();
			const spelling = name + (callee.typeSuffix ?? '');
			if (!TYPEOF_SYNTAX_INTRINSICS.has(spelling) && !TYPEOF_ACCEPTED_COMPARISONS.has(spelling)) { return undefined; }
			if ((scope ??= sourceNameScopeFor(symbols, member, visible)).runtimeShadows.has(name)) { return undefined; }
			if (!externalNames) {
				externalNames = new Set((context?.implicitMembers ?? []).map(item => item.name.toLowerCase()));
				for (const type of context?.projectClassMembers ?? []) {
					if (type.kind === 'document' || type.kind === 'userform' || type.kind === 'class' && type.predeclaredId !== false) {
						externalNames.add(type.name.toLowerCase());
					}
				}
			}
			return externalNames.has(name) ? undefined : spelling;
		};
	};
}

function measuredIsSpan(
	expr: ExprNode,
	intrinsic: string | undefined,
	shapeOf: ReturnType<ReturnType<typeof expressionShapeResolver>>,
	member: ProcedureNode,
): Span {
	if (expr.exprKind === 'IndexExpr') {
		const argument = expr.args[0]?.value;
		// Numeric Int/Fix and CBool select the argument; the measured string
		// Int/Fix forms select the name. Source functions retain callable spans.
		if (argument && (intrinsic === 'cbool' || intrinsic === 'cint' || intrinsic === 'clng' || intrinsic === 'int' || intrinsic === 'fix') && numericIsOperand(argument, shapeOf)) {
			return argument.span;
		}
		if (measuredComparisonCall(expr, intrinsic)) {
			const commaIndex = intrinsic === 'instrb' ? 1 : 0;
			const second = expr.args[commaIndex].value!;
			const third = expr.args[commaIndex + 1].value!;
			const gap = retainedExpressionText(member.body, { start: second.span.end, end: third.span.start });
			const comma = gap?.match(/^\s*,/);
			if (comma) { return { start: expr.span.start, end: second.span.end + comma[0].length }; }
		}
	}
	return compileOperandSpan(expr, member);
}

// Exact positional arities observed in the three retained batches. Unmeasured
// named/other-arity forms retain the ordinary conservative shape behavior.
const TYPEOF_ACCEPTED_COMPARISONS = new Map([['instrb', 3], ['instr', 2], ['strcomp', 2]]);
function measuredComparisonCall(expr: ExprNode, name: string | undefined): boolean {
	return name !== undefined && expr.exprKind === 'IndexExpr'
		&& expr.args.length === TYPEOF_ACCEPTED_COMPARISONS.get(name)
		&& expr.args.every(arg => !arg.name && arg.value);
}
