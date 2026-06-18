// Detects control flow that the structural If/ElseIf/Else branch-merge does not
// model soundly, so dataflow rules can fall back to the conservative
// straight-line walk for such procedures.

import { statementTokens, tokensWithoutLeadingLineNumber, tokenWord } from '../lexer/tokenHelpers';
import type { BodyNode, LeafStatementNode, ProcedureNode, Span } from '../parser/nodes';
import { isLeafStatement } from '../parser/nodes';
import type { ConditionalActivityTracker } from '../conditional/conditionalCompilation';
import {
	collectProcedureLabelDeclarations,
	collectProcedureLabelReferences,
} from './procedureLabels';

/**
 * True when a procedure contains control flow that can skip or re-run
 * assignments in ways the structural branch-merge cannot see: any label, any
 * GoTo / GoSub / On..GoTo / On..GoSub / Resume target, or any `On Error` /
 * `Resume` statement (whose exception edges can bypass an assignment that the
 * merge would otherwise assume ran on a branch). Such procedures fall back to the
 * conservative straight-line dataflow (blanket demotion), preserving the no-FP
 * contract.
 */
// Per-parse memo: the result is a pure function of (source, procedure, activity),
// and within one analysis pass a given procedure node is always paired with the
// same source/activity (and is a fresh node on the next parse), so keying on the
// node matches the engine's per-pass WeakMap convention. Both dataflow rules that
// gate on this share the cached boolean instead of each re-walking the body (up to
// three walks per call).
const UNSTRUCTURED_FLOW_CACHE = new WeakMap<ProcedureNode, boolean>();

export function procedureHasUnstructuredFlow(
	source: string,
	procedure: ProcedureNode,
	activity?: ConditionalActivityTracker,
): boolean {
	const cached = UNSTRUCTURED_FLOW_CACHE.get(procedure);
	if (cached !== undefined) {
		return cached;
	}
	const result = computeProcedureHasUnstructuredFlow(source, procedure, activity);
	UNSTRUCTURED_FLOW_CACHE.set(procedure, result);
	return result;
}

function computeProcedureHasUnstructuredFlow(
	source: string,
	procedure: ProcedureNode,
	activity?: ConditionalActivityTracker,
): boolean {
	if (collectProcedureLabelReferences(source, procedure, activity).length > 0) {
		return true;
	}
	if (collectProcedureLabelDeclarations(source, procedure, activity).length > 0) {
		return true;
	}
	// `On Error Resume Next` / `On Error GoTo 0` / bare `Resume` / `Resume Next`
	// carry no label, so the collectors above miss them — scan for them directly.
	return hasOnErrorOrResumeStatement(procedure.body, source, activity);
}

function hasOnErrorOrResumeStatement(
	body: readonly BodyNode[],
	source: string,
	activity?: ConditionalActivityTracker,
): boolean {
	for (const node of body) {
		if (activity?.isInactive(node.span)) {
			continue;
		}
		if (isLeafStatement(node)) {
			if (isOnErrorOrResume(source, node.span)) {
				return true;
			}
		} else if ('body' in node && Array.isArray(node.body)) {
			if (hasOnErrorOrResumeStatement(node.body, source, activity)) {
				return true;
			}
		}
	}
	return false;
}

function isOnErrorOrResume(source: string, span: Span): boolean {
	const toks = tokensWithoutLeadingLineNumber(statementTokens(source, span));
	if (toks.length === 0) {
		return false;
	}
	const first = tokenWord(toks[0]);
	if (first === 'resume') {
		return true;
	}
	return first === 'on' && tokenWord(toks[1]) === 'error';
}
