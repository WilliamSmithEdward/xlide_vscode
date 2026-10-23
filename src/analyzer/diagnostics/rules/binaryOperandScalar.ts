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
// coerce). Parentheses keep the aggregate, `(a) + 1` (issue #1415). Everything
// else stays quiet - scalars, Variant, Object/host receivers, an indexed element
// `a(i)` (an `IndexExpr`, not the bare aggregate), member access, calls,
// cross-module UDTs (deferred), and undeclared names. `Is` (object operands,
// owned by `is-operator-non-object`) is out of scope; `Like` refuses an array,
// and its Type operand is `udt-value-mismatch`'s (issue #1415).
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
import { caseComparisonOperator } from '../statementOperands';

/**
 * Binary operators that require scalar operands (MS-VBAL 5.6). Excludes `Is`
 * (object-reference operands, owned by `is-operator-non-object`), `Like`
 * (operands are coerced to String) and `^`: the VBE compiles `a ^ 2` on an
 * array or a Type, which runs or raises 51 (measured in Excel 16.0).
 */
const SCALAR_OPERAND_OPERATORS: ReadonlySet<BinaryOperator> = new Set<BinaryOperator>([
	'&', '+', '-', '*', '/', '\\', 'Mod',
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
			// Nothing is no value for any operator but Is: `Nothing + 1`,
			// `"a" & Nothing`, `-Nothing`, `Not Nothing` (issue #325, measured
			// in Excel 16.0).
			const operation = expr.exprKind === 'BinaryExpr' && expr.operator !== 'Is'
				? { operator: expr.operator, nothing: [expr.left, expr.right].find(isNothing) }
				: expr.exprKind === 'UnaryExpr' ? { operator: expr.operator, nothing: isNothing(expr.operand) ? expr.operand : undefined } : undefined;
			if (operation?.nothing) {
				push(
					'nonScalarBinaryOperand',
					`The '${operation.operator}' operator needs a value, and Nothing is an object reference. This is a VBE compile error: Invalid use of object.`,
					operation.nothing.span,
				);
				return;
			}
			// A unary minus negates one value, and an array is none: `-a` is
			// Type mismatch, where `Not a` compiles (measured in Excel 16.0).
			if (expr.exprKind === 'UnaryExpr' && expr.operator === '-' && expr.operand.exprKind === 'IdentifierExpr'
				&& shapes.get(expr.operand.name.toLowerCase())?.isArray === true) {
				push(
					'nonScalarBinaryOperand',
					`The '-' operator requires a scalar operand, but '${expr.operand.name}' is declared As an array. This will fail to compile with 'Type mismatch'.`,
					expr.operand.span,
				);
				return;
			}
			// `Case a` compares the Select value with a, and an array is no value
			// to compare (issue #1415, measured in Excel, Word, PowerPoint and
			// Access 16.0).
			const comparison = caseComparisonOperator(expr);
			const comparand = comparison ? nonScalarOperand(expr, shapes, udtNames, true) : undefined;
			if (comparand) {
				push(
					'nonScalarBinaryOperand',
					`The Case's '${comparison}' comparison requires a scalar operand, but ${comparand.detail}. This will fail to compile with 'Type mismatch'.`,
					comparand.span,
				);
				return;
			}
			// `a Like "*"` and `"x" Like a` refuse an array too, where Like
			// takes any value as a String; a Type there is udt-value-mismatch's
			// (issue #1415, measured in Excel, Word, PowerPoint and Access 16.0).
			const like = expr.exprKind === 'BinaryExpr' && expr.operator === 'Like';
			if (expr.exprKind !== 'BinaryExpr' || (!SCALAR_OPERAND_OPERATORS.has(expr.operator) && !like)) {
				return;
			}
			const offender = nonScalarOperand(expr.left, shapes, udtNames, like)
				?? nonScalarOperand(expr.right, shapes, udtNames, like);
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

function isNothing(expr: ExprNode): boolean {
	const inner = expr.exprKind === 'ParenExpr' ? expr.inner : expr;
	return inner.exprKind === 'LiteralExpr' && inner.literalKind === 'nothing';
}

/** A provably non-scalar identifier operand (a bare array, or a same-module
 *  user-defined Type value), or undefined. Quiet for everything else. */
function nonScalarOperand(
	operand: ExprNode,
	shapes: ReadonlyMap<string, DeclaredValueShape>,
	udtNames: ReadonlySet<string>,
	arraysOnly = false,
): { span: Span; detail: string } | undefined {
	// Parentheses keep the aggregate: `(a) + 1` and `((a)) & 1` are Type
	// mismatch too (issue #1415, measured in Excel, Word, PowerPoint and
	// Access 16.0).
	let expr = operand;
	while (expr.exprKind === 'ParenExpr') {
		expr = expr.inner;
	}
	// Only a bare identifier is typed: an indexed element `a(i)` is an IndexExpr
	// (a scalar element), member access / calls are not the bare aggregate, and
	// literals are scalar.
	if (expr.exprKind !== 'IdentifierExpr') {
		return undefined;
	}
	const shape = shapes.get(expr.name.toLowerCase());
	if (shape?.isArray === true) {
		return { span: expr.span, detail: `'${expr.name}' is declared As an array` };
	}
	const declared = shape?.asType;
	if (!arraysOnly && declared && udtNames.has(declared.toLowerCase())) {
		return {
			span: expr.span,
			detail: `'${expr.name}' is declared As ${declared} (a user-defined Type)`,
		};
	}
	return undefined; // scalar / Variant / Object / class / cross-module UDT / undeclared -> quiet
}
