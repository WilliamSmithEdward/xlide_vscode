// Scalar-requiring operators reject declared arrays and same-module UDTs,
// including parentheses, field chains and known call results. Unknown names,
// Variant and object/default-member values stay quiet. Excel accepts aggregate
// operands of ^ (measured 2026-10-15), so that operator is deliberately excluded.

import type {
	BinaryOperator,
	BodyNode,
	ExprNode,
	Span,
} from '../../parser/nodes';
import type { VbaSymbol } from '../../symbols/symbolModel';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { PushFn } from '../analysisContext';
import {
	sameModuleTypeNames,
	type DeclaredValueShape,
} from '../typeInference';
import { expressionShapeResolver, type ExpressionShapeContext } from '../expressionShape';
import type { ProcedureExpressionVisitor } from '../exprWalk';

/**
 * Binary operators that require scalar operands (MS-VBAL 5.6). Excludes `Is`
 * (object-reference operands, owned by `is-operator-non-object`) and `^`:
 * the VBE compiles `a ^ 2` on an array or a Type, which runs or raises 51,
 * where `Like` on one is Type mismatch (measured in Excel 16.0).
 */
const SCALAR_OPERAND_OPERATORS: ReadonlySet<BinaryOperator> = new Set<BinaryOperator>([
	'&', '+', '-', '*', '/', '\\', 'Mod',
	'=', '<>', '<', '>', '<=', '>=',
	'And', 'Or', 'Xor', 'Eqv', 'Imp', 'Like',
]);

export function checkBinaryOperandScalar(
	symbols: ReturnType<typeof buildModuleSymbols>,
	push: PushFn,
	projectVisibleSymbols?: readonly VbaSymbol[],
	shapeContext?: ExpressionShapeContext,
): ProcedureExpressionVisitor {
	const udtNames = sameModuleTypeNames(symbols);
	const forProcedure = expressionShapeResolver(symbols, projectVisibleSymbols, shapeContext);
	return (member) => {
		const shapeOf = forProcedure(member);
		return (expr) => {
			// Nothing is no value for any operator but Is: `Nothing + 1`,
			// `"a" & Nothing`, `-Nothing`, `Not Nothing` (issue #325, measured
			// in Excel 16.0).
			const operator = expr.exprKind === 'BinaryExpr' || expr.exprKind === 'UnaryExpr' ? expr.operator : undefined;
			const nothing = expr.exprKind === 'BinaryExpr' && expr.operator !== 'Is'
				? (isNothing(expr.left) ? expr.left : isNothing(expr.right) ? expr.right : undefined)
				: expr.exprKind === 'UnaryExpr' && isNothing(expr.operand) ? expr.operand : undefined;
			if (nothing) {
				push(
					'nonScalarBinaryOperand',
					`The '${operator}' operator needs a value, and Nothing is an object reference. This is a VBE compile error: Invalid use of object.`,
					nothing.span,
				);
				return;
			}
			// A unary minus negates one value, and an array is none: `-a` is
			// Type mismatch, where `Not a` compiles (measured in Excel 16.0).
			if (expr.exprKind === 'UnaryExpr' && expr.operator === '-' && shapeOf(expr.operand)?.isArray === true) {
				const operand = nonScalarOperand(expr.operand, shapeOf, udtNames)!;
				push(
					'nonScalarBinaryOperand',
					`The '-' operator requires a scalar operand, but ${operand.detail}. This will fail to compile with 'Type mismatch'.`,
					operand.span,
				);
				return;
			}
			if (expr.exprKind !== 'BinaryExpr' || !SCALAR_OPERAND_OPERATORS.has(expr.operator)) {
				return;
			}
			const offender = nonScalarOperand(expr.left, shapeOf, udtNames)
				?? nonScalarOperand(expr.right, shapeOf, udtNames);
			if (offender) {
				push(
					'nonScalarBinaryOperand',
					`The '${expr.operator}' operator requires a scalar operand, but ${offender.detail}. This will fail to compile with 'Type mismatch'.`,
					// Measured file-slot array addition selects the operator. Keep
					// other operators/UDTs on their established operand spans.
					expr.operator === '+' && shapeOf(expr.left)?.isArray
						? additionOperatorSpan(member.body, expr) ?? offender.span : offender.span,
				);
			}
		};
	};
}

function isNothing(expr: ExprNode): boolean {
	while (expr.exprKind === 'ParenExpr') { expr = expr.inner; }
	return expr.exprKind === 'LiteralExpr' && expr.literalKind === 'nothing';
}

/** A provably non-scalar operand; unknown shapes stay quiet. */
function nonScalarOperand(
	expr: ExprNode,
	shapeOf: (expr: ExprNode) => DeclaredValueShape | undefined,
	udtNames: ReadonlySet<string>,
): { span: Span; detail: string } | undefined {
	const shape = shapeOf(expr);
	if (shape?.isArray === true) {
		return { span: expr.span, detail: `${expr.exprKind === 'IdentifierExpr' ? `'${expr.name}'` : 'the expression'} has array type` };
	}
	const declared = shape?.asType;
	if (declared && udtNames.has(declared.toLowerCase())) {
		return {
			span: expr.span,
			detail: `${expr.exprKind === 'IdentifierExpr' ? `'${expr.name}'` : 'the expression'} has type ${declared} (a user-defined Type)`,
		};
	}
	return undefined; // scalar / Variant / Object / class / cross-module UDT / undeclared -> quiet
}

/** Read only the containing raw statement, using ordered AST spans. This avoids
 * a backwards source/module scan for each diagnostic. Structured expressions
 * without retained source deliberately return undefined. */
export function retainedExpressionText(body: readonly BodyNode[], span: Span): string | undefined {
	let low = 0;
	let high = body.length;
	while (low < high) {
		const mid = (low + high) >>> 1;
		if (body[mid].span.end <= span.start) { low = mid + 1; } else { high = mid; }
	}
	const node = body[low];
	if (!node || node.span.start > span.start || node.span.end < span.end) { return undefined; }
	if (node.kind === 'Statement') { return node.raw.slice(span.start - node.span.start, span.end - node.span.start); }
	if ('body' in node) { return retainedExpressionText(node.body, span); }
	return undefined;
}

function additionOperatorSpan(body: readonly BodyNode[], expr: Extract<ExprNode, { exprKind: 'BinaryExpr' }>): Span | undefined {
	const gap = { start: expr.left.span.end, end: expr.right.span.start };
	const text = retainedExpressionText(body, gap);
	// Whitespace and explicit line continuations can separate the operands.
	const match = text?.match(/^(?:\s|_\r?\n)*([+])(?:\s|_\r?\n)*$/);
	if (!match) { return undefined; }
	const start = gap.start + text!.indexOf('+');
	return { start, end: start + 1 };
}
