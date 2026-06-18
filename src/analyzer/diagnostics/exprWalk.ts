// Shared expression-tree traversal for diagnostics that consume the §5.6
// expression AST (MS-VBAL 5.6).
//
// Several rules walk a procedure body to visit specific expression nodes —
// `TypeOf ... Is`, binary `Is`, scalar-requiring binary operators, and (as the
// binder cashes in more families) others. They all need the SAME two-level walk:
// find every root expression in the body (assignment sides, call callee and
// arguments, `If`/`ElseIf` conditions, and nested block bodies), then recurse
// each expression into all of its sub-expressions. Re-implementing that walk per
// rule triplicated a traversal that must stay consistent: a new `BodyNode` or
// `ExprNode` kind would otherwise have to be added in every copy or rules would
// silently miss nodes. This module owns the one walk; rules filter by `exprKind`.

import type { BodyNode, ExprNode } from '../parser/nodes';
import type { ConditionalActivityTracker } from '../conditional/conditionalCompilation';
import { isInactiveNode } from './walker';

/**
 * Visits every expression node — each root expression and all of its
 * sub-expressions — reachable in a procedure body, skipping inactive
 * conditional-compilation regions. Rules narrow to the nodes they care about by
 * testing `expr.exprKind` inside `visit`.
 */
export function forEachExpressionInBody(
	body: readonly BodyNode[],
	activity: ConditionalActivityTracker | undefined,
	visit: (expr: ExprNode) => void,
): void {
	for (const node of body) {
		if (isInactiveNode(activity, node)) {
			continue;
		}
		switch (node.kind) {
			case 'Assignment':
				forEachSubExpression(node.lhs, visit);
				forEachSubExpression(node.rhs, visit);
				break;
			case 'Call':
				forEachSubExpression(node.callee, visit);
				for (const arg of node.args) {
					if (arg.value) {
						forEachSubExpression(arg.value, visit);
					}
				}
				break;
			case 'IfBlock':
				for (const branch of node.branches) {
					if (branch.condition) {
						forEachSubExpression(branch.condition, visit);
					}
				}
				// Arm statements live in the flat body; recurse it for nested exprs.
				forEachExpressionInBody(node.body, activity, visit);
				break;
			default:
				if ('body' in node && Array.isArray(node.body)) {
					forEachExpressionInBody(node.body, activity, visit);
				}
		}
	}
}

/** Visits `expr` and every nested sub-expression (pre-order). */
export function forEachSubExpression(expr: ExprNode, visit: (expr: ExprNode) => void): void {
	visit(expr);
	switch (expr.exprKind) {
		case 'BinaryExpr':
			forEachSubExpression(expr.left, visit);
			forEachSubExpression(expr.right, visit);
			break;
		case 'UnaryExpr':
			forEachSubExpression(expr.operand, visit);
			break;
		case 'ParenExpr':
			forEachSubExpression(expr.inner, visit);
			break;
		case 'IndexExpr':
			forEachSubExpression(expr.callee, visit);
			for (const arg of expr.args) {
				if (arg.value) {
					forEachSubExpression(arg.value, visit);
				}
			}
			break;
		case 'MemberAccessExpr':
			if (expr.object) {
				forEachSubExpression(expr.object, visit);
			}
			break;
		case 'TypeOfIsExpr':
			forEachSubExpression(expr.operand, visit);
			break;
		default:
			break; // LiteralExpr / IdentifierExpr / NewExpr / AddressOfExpr: leaves
	}
}
