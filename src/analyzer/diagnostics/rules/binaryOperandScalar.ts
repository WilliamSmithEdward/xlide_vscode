// Non-scalar operand of a scalar-requiring binary operator (XLIDE v2.5.0).
//
// A bare array variable, or a same-module user-defined `Type` (struct) value,
// used as an operand of a binary operator that requires a scalar - `&`
// concatenation, arithmetic (`+ - * / \ ^ Mod`), comparison (`= <> < > <= >=`),
// or Boolean/bitwise (`And Or Xor Eqv Imp`) - is a VBE compile error ("Type
// mismatch"). Oracle-verified rejected at COMPILE across all four operator
// classes for both an array operand (`nonscalar_array_concat_probe`,
// `_arith_plus_probe`, `_comparison_lt_probe`, `_bool_and_probe`) and a UDT
// operand (`nonscalar_udt_concat_probe`, `_comparison_lt_probe`,
// `_arith_plus_probe`), with scalar / Variant / indexed element `a(0)` / Excel
// Range operands accepted as controls.
//
// No-false-positive discipline (mirrors `is-operator-non-object`): fires only on a
// plain `IdentifierExpr` operand that is PROVABLY non-scalar - its declaration
// shape is an array (`DeclaredValueShape.isArray`), or its declared type names a
// `Type` declared in this module (`Type` names are always user identifiers, never
// scalar keywords, so this never collides with a scalar; an Object/class instance
// is not a module `Type` symbol and stays quiet, since a default member may
// coerce). Everything else stays quiet - scalars, Variant, Object/host receivers,
// an indexed element `a(i)` (an `IndexExpr`, not the bare aggregate), member
// access, calls, parenthesised expressions, cross-module UDTs (deferred), and
// undeclared names. `Is` (object operands, owned by `is-operator-non-object`) and
// `Like` (operands coerced to String) are out of scope.
//
// The body/expression traversal is the shared `forEachExpressionInBody` walk
// (diagnostics/exprWalk.ts); this rule only filters to scalar-requiring binaries.

import type {
	BinaryOperator,
	ExprNode,
	Span,
} from '../../parser/nodes';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { PushFn } from '../analysisContext';
import {
	declarationShapeEnvironmentFor,
	sameModuleTypeNames,
	type DeclaredValueShape,
} from '../typeInference';
import type { ProcedureExpressionVisitor } from '../exprWalk';

/**
 * Binary operators that require scalar operands (MS-VBAL 5.6). Excludes `Is`
 * (object-reference operands, owned by `is-operator-non-object`) and `Like`
 * (operands are coerced to String).
 */
const SCALAR_OPERAND_OPERATORS: ReadonlySet<BinaryOperator> = new Set<BinaryOperator>([
	'&', '+', '-', '*', '/', '\\', '^', 'Mod',
	'=', '<>', '<', '>', '<=', '>=',
	'And', 'Or', 'Xor', 'Eqv', 'Imp',
]);

export function checkBinaryOperandScalar(
	symbols: ReturnType<typeof buildModuleSymbols>,
	push: PushFn,
): ProcedureExpressionVisitor {
	const udtNames = sameModuleTypeNames(symbols);
	return (member) => {
		const shapes = declarationShapeEnvironmentFor(symbols, member);
		return (expr) => {
			if (expr.exprKind !== 'BinaryExpr' || !SCALAR_OPERAND_OPERATORS.has(expr.operator)) {
				return;
			}
			const offender = nonScalarOperand(expr.left, shapes, udtNames)
				?? nonScalarOperand(expr.right, shapes, udtNames);
			if (offender) {
				push(
					'nonScalarBinaryOperand',
					`The '${expr.operator}' operator requires a scalar operand, but ${offender.detail}. This will fail to compile with 'Type mismatch'.`,
					offender.span,
				);
			}
		};
	};
}

/** A provably non-scalar identifier operand (a bare array, or a same-module
 *  user-defined Type value), or undefined. Quiet for everything else. */
function nonScalarOperand(
	expr: ExprNode,
	shapes: ReadonlyMap<string, DeclaredValueShape>,
	udtNames: ReadonlySet<string>,
): { span: Span; detail: string } | undefined {
	// Only a bare identifier is typed: an indexed element `a(i)` is an IndexExpr
	// (a scalar element), member access / calls / parens are not the bare
	// aggregate, and literals are scalar.
	if (expr.exprKind !== 'IdentifierExpr') {
		return undefined;
	}
	const shape = shapes.get(expr.name.toLowerCase());
	if (shape?.isArray === true) {
		return { span: expr.span, detail: `'${expr.name}' is declared As an array` };
	}
	const declared = shape?.asType;
	if (declared && udtNames.has(declared.toLowerCase())) {
		return {
			span: expr.span,
			detail: `'${expr.name}' is declared As ${declared} (a user-defined Type)`,
		};
	}
	return undefined; // scalar / Variant / Object / class / cross-module UDT / undeclared -> quiet
}
