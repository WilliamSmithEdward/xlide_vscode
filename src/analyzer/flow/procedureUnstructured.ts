// Detects control flow that the structural If/ElseIf/Else branch-merge does not
// model soundly, so dataflow rules can fall back to the conservative
// straight-line walk for such procedures.

import type { BodyNode, ProcedureNode } from '../parser/nodes';
import { isLeafStatement } from '../parser/nodes';
import type { ConditionalActivityTracker } from '../conditional/conditionalCompilation';
import { statementHasUnstructuredFlow } from './procedureLabels';

/**
 * True when a procedure contains control flow that can skip or re-run
 * assignments in ways the structural branch-merge cannot see: any label, any
 * GoTo / GoSub / On..GoTo / On..GoSub / Resume target, or any `On Error` /
 * `Resume` statement (whose exception edges can bypass an assignment that the
 * merge would otherwise assume ran on a branch). Such procedures fall back to the
 * conservative straight-line dataflow (blanket demotion), preserving the no-FP
 * contract.
 */
// Procedure nodes are reused by the parse cache. Reuse flow facts only while
// the source and conditional activity that produced them remain the same.
const UNSTRUCTURED_FLOW_CACHE = new WeakMap<ProcedureNode, {
	source: string;
	activity: ConditionalActivityTracker | undefined;
	result: boolean;
}>();

export function procedureHasUnstructuredFlow(
	source: string,
	procedure: ProcedureNode,
	activity?: ConditionalActivityTracker,
): boolean {
	const cached = UNSTRUCTURED_FLOW_CACHE.get(procedure);
	if (cached && cached.source === source && cached.activity === activity) {
		return cached.result;
	}
	const result = hasUnstructuredStatement(procedure.body, source, activity);
	UNSTRUCTURED_FLOW_CACHE.set(procedure, { source, activity, result });
	return result;
}

function hasUnstructuredStatement(
	body: readonly BodyNode[],
	source: string,
	activity?: ConditionalActivityTracker,
): boolean {
	for (const node of body) {
		if (activity?.isInactive(node.span)) {
			continue;
		}
		if (isLeafStatement(node)) {
			if (statementHasUnstructuredFlow(source, node.span)) {
				return true;
			}
		} else if ('body' in node && Array.isArray(node.body)) {
			if (hasUnstructuredStatement(node.body, source, activity)) {
				return true;
			}
		}
	}
	return false;
}
