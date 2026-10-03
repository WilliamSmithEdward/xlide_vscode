import type { ConditionalActivityTracker } from '../conditional/conditionalCompilation';
import type { BodyNode, LeafStatementNode } from './nodes';
import { isLeafStatement } from './nodes';

/** Walks every leaf statement (Assignment/Call/Statement) in a body, descending into nested blocks. */
export function forEachStatement(
	body: BodyNode[],
	visit: (stmt: LeafStatementNode) => void,
	activity?: ConditionalActivityTracker,
): void {
	for (const node of body) {
		if (activity?.isInactive(node.span)) {
			continue;
		}
		if (isLeafStatement(node)) {
			visit(node);
		} else if ('body' in node && Array.isArray(node.body)) {
			forEachStatement(node.body, visit, activity);
		}
	}
}
