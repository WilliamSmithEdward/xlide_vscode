// `TypeOf <expr> Is <Type>` always-False detection (MS-VBAL 5.6 TypeOf...Is).
//
// `TypeOf x Is Y` tests whether x's runtime object is-a Y. When x is declared as
// a known, concrete object type A that can never hold a Y, the test is provably
// always False — a dead branch and almost always a bug.
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
	ModuleNode,
	Span,
	TypeOfIsExpr,
} from '../../parser/nodes';
import type { MemberCompletionContext } from '../../completion/memberAccess';
import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { PushFn } from '../analysisContext';
import {
	isKnownScalarType,
	objectAssignmentIncompatibilityReason,
	resolveKnownObjectAssignmentType,
	typeEnvironmentFor,
} from '../typeInference';
import { activeModuleMembers } from '../walker';
import { tokenizeCached } from '../../lexer/tokenize';
import { forEachExpressionInBody } from '../exprWalk';

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

export function checkTypeOfIsCompatibility(
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	memberCtx: MemberCompletionContext,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const env = typeEnvironmentFor(symbols, member);
		forEachExpressionInBody(member.body, activity, (expr) => {
			if (expr.exprKind === 'TypeOfIsExpr') {
				checkTypeOfIs(expr, env, memberCtx, push);
			}
		});
	}
}

function checkTypeOfIs(
	expr: TypeOfIsExpr,
	env: Map<string, string>,
	memberCtx: MemberCompletionContext,
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
	const operandType = resolveKnownObjectAssignmentType(declared, memberCtx);
	const targetType = resolveKnownObjectAssignmentType(expr.typeName, memberCtx);
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
	if (operandType.kind === 'project' && isImplementedByAnyProjectClass(operandType, memberCtx)) {
		return;
	}
	// Mutual incompatibility: neither type is assignable to the other (reuses the
	// existing object tables, which already honour `Implements`). If either is
	// assignable, `TypeOf` could be True, so stay quiet.
	const operandCanBeTarget = objectAssignmentIncompatibilityReason(
		expr.typeName,
		{ type: declared, label: declared, span: expr.span },
		memberCtx,
	) === undefined;
	const targetCanBeOperand = objectAssignmentIncompatibilityReason(
		declared,
		{ type: expr.typeName, label: expr.typeName, span: expr.span },
		memberCtx,
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

/** True when any project class declares `Implements <operandType>`. */
function isImplementedByAnyProjectClass(
	operandType: Extract<ReturnType<typeof resolveKnownObjectAssignmentType>, { kind: 'project' }>,
	memberCtx: MemberCompletionContext,
): boolean {
	const names = new Set([operandType.key, operandType.display.toLowerCase()]);
	return (memberCtx.projectClassMembers ?? []).some((projectType) =>
		(projectType.implements ?? []).some((implemented) => names.has(implemented.toLowerCase())),
	);
}

/**
 * Rule: the binary `Is` object-identity operator requires object-reference
 * operands on both sides (MS-VBAL 5.6). When an operand is PROVABLY a non-object
 * — a scalar value literal, or an identifier declared As a known scalar type
 * (numeric / String / Boolean / Date) — `a Is b` is a VBE compile error
 * ("Object required"). Oracle-verified rejected at compile:
 * `is_scalar_long_var_compile`, `is_scalar_string_var_compile`,
 * `is_two_scalar_vars_compile`, `is_literal_integer_operands_probe`,
 * `is_literal_string_operands_probe`; with the Variant control
 * `is_variant_var_ok_compile` and the Object control `is_object_var_vs_nothing_probe`
 * accepted.
 *
 * No-false-positive discipline: fires only when an operand is PROVABLY scalar.
 * Variant (can hold an object reference), Object/class-typed operands,
 * `Nothing`/`Null`/`Empty` literals, member access, call results, arrays, and any
 * unresolved / undeclared identifier all stay quiet. `TypeOf x Is Y` is a
 * distinct `TypeOfIsExpr`, not a binary `Is`, so it is unaffected.
 */
export function checkIsOperatorOperands(
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const env = typeEnvironmentFor(symbols, member);
		forEachExpressionInBody(member.body, activity, (expr) => {
			if (expr.exprKind !== 'BinaryExpr' || expr.operator !== 'Is') {
				return;
			}
			const offender = nonObjectOperand(expr.left, env) ?? nonObjectOperand(expr.right, env);
			if (offender) {
				push(
					'isOperatorNonObject',
					`The 'Is' operator requires object operands, but ${offender.detail}, which is not an object.`,
					offender.span,
				);
			}
		});
	}
}

/** Describes a provably non-object (scalar) operand of `Is`, or undefined. */
function nonObjectOperand(
	expr: ExprNode,
	env: Map<string, string>,
): { span: Span; detail: string } | undefined {
	if (expr.exprKind === 'LiteralExpr') {
		const k = expr.literalKind;
		if (k === 'integer' || k === 'float' || k === 'string' || k === 'date' || k === 'boolean') {
			return { span: expr.span, detail: `'${expr.raw}' is a ${k} literal` };
		}
		return undefined; // Nothing / Null / Empty -> not provably scalar
	}
	if (expr.exprKind === 'IdentifierExpr') {
		const declared = env.get(expr.name.toLowerCase());
		if (!declared) {
			return undefined; // undeclared / unknown -> quiet
		}
		// Match the declared type name against the built-in scalar set WITHOUT
		// normalizeType's leading-`vb` strip: a user class named `vbLong`/`vbString`
		// is a real object type that the strip would collapse to a scalar word and
		// wrongly flag (adversarial FP-hunt finding; VBE compiles `x As vbLong Is
		// Nothing`). Strip only a trailing array `()` marker.
		const raw = declared.replace(/\s*\(\s*\)\s*$/, '').trim().toLowerCase();
		if (isKnownScalarType(raw)) {
			return { span: expr.span, detail: `'${expr.name}' is declared As ${declared}` };
		}
		return undefined; // Variant / Object / class -> quiet
	}
	return undefined; // member / call / paren / array / New / unary -> quiet (v1)
}
