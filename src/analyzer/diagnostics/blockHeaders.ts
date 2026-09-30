// A block's own lines as statements (issues #233 and #237): the line a For,
// Select, Do, While or With opens with, a Do's `Loop While` line, and each
// If arm's condition line. Rules judge the expressions there, and a walk that
// enters a block counts the names they pass ByRef as ones the block changes.

import { tokenizeCached } from '../lexer/tokenize';
import { statementTokensCached, tokenWord as tokenText } from '../lexer/tokenHelpers';
import type { VbaToken } from '../lexer/tokenKinds';
import type { BodyNode, DoBlockNode, StatementNode } from '../parser/nodes';

/**
 * The block kinds whose own line evaluates an expression: a For's bounds and
 * step, a Select Case subject, a Do or While condition, a With subject.
 */
const HEADER_BLOCKS: ReadonlySet<string> = new Set(['ForBlock', 'SelectBlock', 'DoBlock', 'WhileBlock', 'WithBlock']);

/**
 * A block's header line as a statement of its own, and a Do's `Loop While`
 * or `Loop Until` line (issue #233). The header ends at the end of its
 * logical line or at a colon, so `If a Then With c: .Add 1: End With` gives
 * `With c` alone, and a string holding a colon stays whole.
 */
export function blockHeaderStatements(source: string, node: BodyNode): { before?: StatementNode; after?: StatementNode } {
	if (!HEADER_BLOCKS.has(node.kind)) {
		return {};
	}
	const toks = tokenizeCached(source);
	const statement = (from: number, to: number): StatementNode => {
		const span = { start: toks[from].start, end: toks[to].end };
		return { kind: 'Statement', span, raw: source.slice(span.start, span.end) };
	};
	const separator = (tok: VbaToken): boolean => tok.kind === 'newline' || tok.kind === 'colon';
	// The first token at or after the block's start.
	let lo = 0;
	let hi = toks.length;
	while (lo < hi) {
		const mid = (lo + hi) >> 1;
		if (toks[mid].start < node.span.start) {
			lo = mid + 1;
		} else {
			hi = mid;
		}
	}
	const out: { before?: StatementNode; after?: StatementNode } = {};
	let end = lo;
	while (end + 1 < toks.length && !separator(toks[end + 1]) && toks[end + 1].kind !== 'comment') {
		end++;
	}
	if (lo < toks.length && !separator(toks[lo])) {
		out.before = statement(lo, end);
	}
	if (node.kind === 'DoBlock' && (node as DoBlockNode).closed) {
		// The last token of the block, and back to the start of its statement.
		let last = hi;
		while (last < toks.length && toks[last].end <= node.span.end) {
			last++;
		}
		last--;
		while (last > end && toks[last].kind === 'comment') {
			last--;
		}
		let first = last;
		while (first - 1 > end && !separator(toks[first - 1])) {
			first--;
		}
		if (first > end && tokenText(toks[first]) === 'loop' && first < last) {
			out.after = statement(first, last);
		}
	}
	return out;
}

/**
 * Every line of a block that can change a variable it names, as a walk that
 * enters the block must know: blockHeaderStatements, and each If and ElseIf
 * condition, `If TryGet(k, obj) Then` among them. A With's subject counts:
 * its body acts on it through `.Add` and the like without naming it.
 */
export function blockHeaderLeaves(source: string, node: BodyNode): StatementNode[] {
	if (node.kind === 'IfBlock') {
		return node.branches
			.filter((branch) => branch.branchKind !== 'else')
			.map((branch) => ({ kind: 'Statement', span: branch.headerSpan, raw: source.slice(branch.headerSpan.start, branch.headerSpan.end) }));
	}
	const { before, after } = blockHeaderStatements(source, node);
	return [before, after].filter((stmt): stmt is StatementNode => stmt !== undefined);
}

/** True for a For, Do or While, whose body may run more than once. */
export function isLoopBlock(node: BodyNode): boolean {
	return node.kind === 'ForBlock' || node.kind === 'DoBlock' || node.kind === 'WhileBlock';
}

/**
 * A Select block's body as its arms: each `Case` line with the statements
 * under it. Only one arm runs, so a walk enters each from the same state.
 */
export function selectArms(source: string, body: readonly BodyNode[]): BodyNode[][] {
	const arms: BodyNode[][] = [];
	for (const node of body) {
		const head = node.kind === 'Statement' ? tokenText(statementTokensCached(source, node.span).find((tok) => tok.kind !== 'integerLiteral')) : '';
		if (head === 'case' || arms.length === 0) {
			arms.push([]);
		}
		arms[arms.length - 1].push(node);
	}
	return arms;
}
