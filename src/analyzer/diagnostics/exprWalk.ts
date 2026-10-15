// Shared expression-tree traversal for diagnostics that consume the §5.6
// expression AST (MS-VBAL 5.6).
//
// Several rules walk a procedure body to visit specific expression nodes -
// `TypeOf ... Is`, binary `Is`, scalar-requiring binary operators, and (as the
// binder cashes in more families) others. They all need the SAME two-level walk:
// find every root expression in the body (assignment sides, call callee and
// arguments, `If`/`ElseIf` conditions, and nested block bodies), then recurse
// each expression into all of its sub-expressions. Re-implementing that walk per
// rule triplicated a traversal that must stay consistent: a new `BodyNode` or
// `ExprNode` kind would otherwise have to be added in every copy or rules would
// silently miss nodes. This module owns the one walk; rules filter by `exprKind`.

import type { BodyNode, ExprNode, LeafStatementNode, ModuleNode, ProcedureNode } from '../parser/nodes';
import type { ConditionalActivityTracker } from '../conditional/conditionalCompilation';
import { activeModuleMembers, blockHeaderStatements, firstExecutableTokenIndex, isInactiveNode, statementAndBranchSpans, statementTokens, tokenText } from './walker';
import { ifBranchStatement } from './blockHeaders';
import { parseExpression, parseParenlessArguments } from '../parser/parseExpression';
import type { VbaToken } from '../lexer/tokenKinds';
import { topLevelEqualsIndex } from '../lexer/tokenHelpers';

/**
 * A rule's per-procedure expression visitor: the factory does the rule's
 * per-member setup (type env, shape env, and so on) and returns a callback invoked for
 * every expression node in that member's body.
 */
export type ProcedureExpressionVisitor = (member: ProcedureNode) => (expr: ExprNode) => void;

/**
 * Runs ONE shared expression walk per active procedure, dispatching every
 * expression node to each registered rule visitor - so N expression-consuming
 * rules cost one tree traversal per body instead of N. Each visitor reports
 * through its own buffer, so registry-order output is preserved.
 */
export function walkProcedureExpressions(
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	factories: readonly ProcedureExpressionVisitor[],
	source?: string,
	hooks?: {
		beforeMember?: (member: ProcedureNode) => void;
		skipBody?: (member: ProcedureNode) => boolean;
	},
	kinds?: readonly (readonly ExprNode['exprKind'][] | undefined)[],
): void {
	if (factories.length === 0) {
		return;
	}
	// Selection is stable for the pass, so build rule indices once, outside
	// the procedure loop. Generic visitors keep the original all-node contract.
	const always: number[] = [];
	const selected: Partial<Record<ExprNode['exprKind'], number[]>> = {};
	let operatorsOnly = kinds !== undefined;
	for (let i = 0; i < factories.length; i++) {
		const takes = kinds?.[i];
		if (!takes) { always.push(i); operatorsOnly = false; continue; }
		for (const kind of takes) {
			(selected[kind] ??= []).push(i);
			if (kind !== 'BinaryExpr' && kind !== 'UnaryExpr' && kind !== 'TypeOfIsExpr') { operatorsOnly = false; }
		}
	}
	for (const kind of Object.keys(selected) as ExprNode['exprKind'][]) {
		selected[kind] = [...selected[kind]!, ...always].sort((a, b) => a - b);
	}
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		hooks?.beforeMember?.(member);
		// Skipped before its visitors are built, as in walkProcedureStatements.
		if (hooks?.skipBody?.(member)) {
			continue;
		}
		const visitors = factories.map((factory) => factory(member));
		forEachExpressionInBody(member.body, activity, (expr) => {
			for (const index of selected[expr.exprKind] ?? always) {
				visitors[index](expr);
			}
		}, source, undefined, operatorsOnly);
	}
}

/**
 * Visits every expression node - each root expression and all of its
 * sub-expressions - reachable in a procedure body, skipping inactive
 * conditional-compilation regions. Rules narrow to the nodes they care about by
 * testing `expr.exprKind` inside `visit`.
 */
export function forEachExpressionInBody(
	body: readonly BodyNode[],
	activity: ConditionalActivityTracker | undefined,
	visit: (expr: ExprNode) => void,
	source?: string,
	conditions?: ReadonlyMap<number, ExprNode>,
	operatorsOnly = false,
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
			case 'Statement': {
				const condition = conditions?.get(node.span.start);
				if (condition) { forEachSubExpression(condition, visit); break; }
				if (source !== undefined && (!operatorsOnly || mayContainCheckedExpression(statementTokens(source, node.span)))) {
					for (const expr of statementValueExpressions(source, node)) { forEachSubExpression(expr, visit); }
				}
				break;
			}
			case 'IfBlock': {
				const first = node.branches[0];
				if (first && !isInactiveNode(activity, { span: first.headerSpan })) {
					if (first.condition) { forEachSubExpression(first.condition, visit); }
					else if (source !== undefined) {
						for (const expr of statementValueExpressions(source, ifBranchStatement(source, first))) { forEachSubExpression(expr, visit); }
					}
				}
				let parsed: Map<number, ExprNode> | undefined;
				for (let i = 1; i < node.branches.length; i++) {
					const branch = node.branches[i];
					if (branch.condition) { (parsed ??= new Map()).set(branch.headerSpan.start, branch.condition); }
				}
				forEachExpressionInBody(node.body, activity, visit, source, parsed, operatorsOnly);
				break;
			}
			default:
				if ('body' in node && Array.isArray(node.body)) {
					const { before, after } = source !== undefined ? blockHeaderStatements(source, node) : {};
					if (source !== undefined && before && !isInactiveNode(activity, before) && (!operatorsOnly || mayContainCheckedExpression(statementTokens(source, before.span)))) {
						for (const expr of statementValueExpressions(source, before)) { forEachSubExpression(expr, visit); }
					}
					forEachExpressionInBody(node.body, activity, visit, source, undefined, operatorsOnly);
					if (source !== undefined && after && !isInactiveNode(activity, after) && (!operatorsOnly || mayContainCheckedExpression(statementTokens(source, after.span)))) {
						for (const expr of statementValueExpressions(source, after)) { forEachSubExpression(expr, visit); }
					}
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

/** Value roots of a leaf or a synthetic block header. Syntax delimiters are
 * removed before parsing, never presented as operators. All rules share the
 * statement token cache; structured leaves reuse the parser's existing AST.
 */
const valueRoots = new WeakMap<readonly VbaToken[], ExprNode[]>();

export function statementValueExpressions(source: string, stmt: LeafStatementNode): ExprNode[] {
	if (stmt.kind === 'Assignment') { return [stmt.rhs, ...targetInputs(stmt.lhs)]; }
	if (stmt.kind === 'Call') { return [...targetInputs(stmt.callee), ...stmt.args.flatMap((arg) => arg.value ? [arg.value] : [])]; }
	const key = statementTokens(source, stmt.span);
	const cached = valueRoots.get(key);
	if (cached) { return cached; }
	const roots: ExprNode[] = [];
	for (const span of statementAndBranchSpans(stmt)) {
		const toks = statementTokens(source, span);
		const first = firstExecutableTokenIndex(toks);
		const head = tokenText(toks[first]);
		const parse = (from: number, to = toks.length): void => {
			if (from >= to) { return; }
			const result = parseExpression(toks, from, to, span.start);
			if (result.expr && result.endIndex === to && result.diagnostics.length === 0) { roots.push(result.expr); }
		};
		// Only top-level separators delimit roots; commas and operators inside
		// a call or parentheses belong to that expression.
		const split = (from: number, separators: ReadonlySet<string>): void => {
			let depth = 0, start = from;
			for (let i = from; i < toks.length; i++) {
				const raw = toks[i].rawText;
				if (raw === '(') { depth++; }
				else if (raw === ')') { depth--; }
				else if (depth === 0 && separators.has(tokenText(toks[i]))) { parse(start, i); start = i + 1; }
			}
			parse(start);
		};
		if (head === 'if' || head === 'elseif') {
			const then = toks.findIndex((tok) => tokenText(tok) === 'then');
			if (then > first) { parse(first + 1, then); }
		} else if (head === 'for') {
			const each = tokenText(toks[first + 1]) === 'each';
			const start = each ? toks.findIndex((tok) => tokenText(tok) === 'in') : topLevelEqualsIndex(toks);
			if (start >= 0) { split(start + 1, new Set(each ? [] : ['to', 'step'])); }
		} else if (head === 'select' && tokenText(toks[first + 1]) === 'case') { parse(first + 2); }
		else if (head === 'while' || head === 'with') { parse(first + 1); }
		else if ((head === 'do' || head === 'loop') && ['while', 'until'].includes(tokenText(toks[first + 1]))) { parse(first + 2); }
		else if (head === 'case') {
			if (tokenText(toks[first + 1]) !== 'else') { split(first + 1, new Set([',', 'to'])); }
		} else if (head === 'debug' && toks[first + 1]?.rawText === '.' && tokenText(toks[first + 2]) === 'print') {
			split(first + 3, new Set([',', ';']));
		} else if (!['dim', 'static', 'const', 'public', 'private', 'redim', 'erase', 'exit', 'goto', 'resume', 'on', 'end', 'next', 'wend', 'else', 'loop', 'do'].includes(head)) {
			const eq = topLevelEqualsIndex(toks);
			if (eq >= 0) { parse(eq + 1); }
			else {
				// In an implicit call, a separated '(' starts a parenthesised
				// argument, not a function call: Take ((a)) & "x".
				let calleeEnd = toks.length;
				if (head !== 'call') {
					let depth = 0;
					for (let i = first + 1; i < toks.length; i++) {
						if (toks[i].rawText === '(') {
							if (depth === 0 && toks[i - 1].end < toks[i].start) { calleeEnd = i; break; }
							depth++;
						} else if (toks[i].rawText === ')') { depth--; }
					}
				}
				const callee = parseExpression(toks, head === 'call' ? first + 1 : first, calleeEnd, span.start);
				if (callee.expr && callee.diagnostics.length === 0) {
					if (callee.endIndex === toks.length && callee.expr.exprKind === 'IndexExpr') {
						for (const arg of callee.expr.args) { if (arg.value) { roots.push(arg.value); } }
					} else {
						for (const arg of parseParenlessArguments(toks, callee.endIndex, toks.length, span.start) ?? []) { if (arg.value) { roots.push(arg.value); } }
					}
				}
			}
		}
	}
	valueRoots.set(key, roots);
	return roots;
}

/** Arguments and receivers are values even when the surrounding target is
 * assigned to or invoked as a statement. Its own bare callee/target is not. */
function targetInputs(expr: ExprNode): ExprNode[] {
	if (expr.exprKind === 'MemberAccessExpr') { return expr.object ? [expr.object] : []; }
	if (expr.exprKind === 'IndexExpr') {
		return [...(expr.callee.exprKind === 'IndexExpr' ? [expr.callee] : targetInputs(expr.callee)),
			...expr.args.flatMap((arg) => arg.value ? [arg.value] : [])];
	}
	return [];
}

/** These AST rules consume operators, not bare argument values. Token-based
 * Sub/member/index checks still request value roots lazily for their candidates. */
function mayContainCheckedExpression(tokens: readonly VbaToken[]): boolean {
	for (const token of tokens) {
		if (token.kind === 'operator' || token.kind === 'keyword' && EXPRESSION_WORDS.has(tokenText(token))) { return true; }
	}
	return false;
}
const EXPRESSION_WORDS = new Set(['typeof', 'is', 'not', 'mod', 'like', 'and', 'or', 'xor', 'eqv', 'imp']);
