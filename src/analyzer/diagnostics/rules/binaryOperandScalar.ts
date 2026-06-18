// Non-scalar operand of a scalar-requiring binary operator (XLIDE v2.5.0).
//
// A bare array variable, or a same-module user-defined `Type` (struct) value,
// used as an operand of a binary operator that requires a scalar — `&`
// concatenation, arithmetic (`+ - * / \ ^ Mod`), comparison (`= <> < > <= >=`),
// or Boolean/bitwise (`And Or Xor Eqv Imp`) — is a VBE compile error ("Type
// mismatch"). Oracle-verified rejected at COMPILE across all four operator
// classes for both an array operand (`nonscalar_array_concat_probe`,
// `_arith_plus_probe`, `_comparison_lt_probe`, `_bool_and_probe`) and a UDT
// operand (`nonscalar_udt_concat_probe`, `_comparison_lt_probe`,
// `_arith_plus_probe`), with scalar / Variant / indexed element `a(0)` / Excel
// Range operands accepted as controls.
//
// No-false-positive discipline (mirrors `is-operator-non-object`): fires only on a
// plain `IdentifierExpr` operand that is PROVABLY non-scalar — its declaration
// shape is an array (`DeclaredValueShape.isArray`), or its declared type names a
// `Type` declared in this module (`Type` names are always user identifiers, never
// scalar keywords, so this never collides with a scalar; an Object/class instance
// is not a module `Type` symbol and stays quiet, since a default member may
// coerce). Everything else stays quiet — scalars, Variant, Object/host receivers,
// an indexed element `a(i)` (an `IndexExpr`, not the bare aggregate), member
// access, calls, parenthesised expressions, cross-module UDTs (deferred), and
// undeclared names. `Is` (object operands, owned by `is-operator-non-object`) and
// `Like` (operands coerced to String) are out of scope.
//
// NOTE: the body/expression walkers below mirror those in `typeOfIs.ts`; a shared
// `forEachBinaryExpr` traversal would unify them (tracked follow-up) — kept local
// here to keep this rule's change focused and low-risk.

import type {
	BinaryExpr,
	BinaryOperator,
	BodyNode,
	ExprNode,
	ModuleNode,
	Span,
} from '../../parser/nodes';
import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { PushFn } from '../analysisContext';
import { declarationShapeEnvironmentFor, type DeclaredValueShape } from '../typeInference';
import { activeModuleMembers, isInactiveNode } from '../walker';

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
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const udtNames = userDefinedTypeNames(symbols);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const shapes = declarationShapeEnvironmentFor(symbols, member);
		forEachScalarOperandBinary(member.body, activity, (expr) => {
			const offender = nonScalarOperand(expr.left, shapes, udtNames)
				?? nonScalarOperand(expr.right, shapes, udtNames);
			if (offender) {
				push(
					'nonScalarBinaryOperand',
					`The '${expr.operator}' operator requires a scalar operand, but ${offender.detail}. This will fail to compile with 'Type mismatch'.`,
					offender.span,
				);
			}
		});
	}
}

/** Names of `Type` (struct) declarations in this module, lowercased. */
function userDefinedTypeNames(
	symbols: ReturnType<typeof buildModuleSymbols>,
): ReadonlySet<string> {
	const names = new Set<string>();
	for (const child of symbols.root.children ?? []) {
		if (child.kind === 'type') {
			names.add(child.name.toLowerCase());
		}
	}
	return names;
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

/** Visits every binary expression with a scalar-requiring operator in a body. */
function forEachScalarOperandBinary(
	body: readonly BodyNode[],
	activity: ConditionalActivityTracker | undefined,
	visit: (expr: BinaryExpr) => void,
): void {
	for (const node of body) {
		if (isInactiveNode(activity, node)) {
			continue;
		}
		switch (node.kind) {
			case 'Assignment':
				forEachScalarOperandBinaryInExpr(node.lhs, visit);
				forEachScalarOperandBinaryInExpr(node.rhs, visit);
				break;
			case 'Call':
				forEachScalarOperandBinaryInExpr(node.callee, visit);
				for (const arg of node.args) {
					if (arg.value) {
						forEachScalarOperandBinaryInExpr(arg.value, visit);
					}
				}
				break;
			case 'IfBlock':
				for (const branch of node.branches) {
					if (branch.condition) {
						forEachScalarOperandBinaryInExpr(branch.condition, visit);
					}
				}
				forEachScalarOperandBinary(node.body, activity, visit);
				break;
			default:
				if ('body' in node && Array.isArray(node.body)) {
					forEachScalarOperandBinary(node.body, activity, visit);
				}
		}
	}
}

/** Recurses an expression, visiting every nested scalar-requiring binary expr. */
function forEachScalarOperandBinaryInExpr(
	expr: ExprNode,
	visit: (expr: BinaryExpr) => void,
): void {
	switch (expr.exprKind) {
		case 'BinaryExpr':
			if (SCALAR_OPERAND_OPERATORS.has(expr.operator)) {
				visit(expr);
			}
			forEachScalarOperandBinaryInExpr(expr.left, visit);
			forEachScalarOperandBinaryInExpr(expr.right, visit);
			break;
		case 'UnaryExpr':
			forEachScalarOperandBinaryInExpr(expr.operand, visit);
			break;
		case 'ParenExpr':
			forEachScalarOperandBinaryInExpr(expr.inner, visit);
			break;
		case 'IndexExpr':
			forEachScalarOperandBinaryInExpr(expr.callee, visit);
			for (const arg of expr.args) {
				if (arg.value) {
					forEachScalarOperandBinaryInExpr(arg.value, visit);
				}
			}
			break;
		case 'MemberAccessExpr':
			if (expr.object) {
				forEachScalarOperandBinaryInExpr(expr.object, visit);
			}
			break;
		case 'TypeOfIsExpr':
			forEachScalarOperandBinaryInExpr(expr.operand, visit);
			break;
		default:
			break; // Literal / Identifier / New / AddressOf -> leaves
	}
}
