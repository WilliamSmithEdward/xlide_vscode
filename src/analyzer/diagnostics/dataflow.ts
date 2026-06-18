// Straight-line local-state dataflow shared by diagnostics rules.
//
// The object-variable-not-set and unallocated-dynamic-array rules both track a
// small three-state lattice per procedure local over straight-line statements:
// every tracked local starts in the rule's initial state, moves through
// rule-specific transitions on plain statements, and demotes to 'unknown'
// when the variable may be rebound on a path the rule does not model - passed
// as a bare (potentially ByRef) call argument, or touched anywhere inside a
// nested runtime block. This module owns that shared walk and the
// call-argument escape scan so the escape analysis cannot drift between
// rules; each rule supplies its own transitions and touch detection.

import type { VbaToken } from '../lexer/tokenKinds';
import { tokenName, tokenWord } from '../lexer/tokenHelpers';
import type { BodyNode, IfBlockNode, LeafStatementNode } from '../parser/nodes';
import { isLeafStatement } from '../parser/nodes';

/** Rule-specific hooks driving one straight-line dataflow walk. */
export interface StraightLineDataflowHooks {
	/** Applies one straight-line statement's transitions and diagnostics. */
	onStatement(stmt: LeafStatementNode): void;
	/** Inspects one non-statement node before its body's touch demotion. */
	onBlock?(node: BodyNode): void;
	/** Lowercased tracked names one nested-block statement touches. */
	touchesInStatement(stmt: LeafStatementNode): Iterable<string>;
	/** Demotes one tracked name to the rule's 'unknown' state. */
	demoteToUnknown(lowerName: string): void;

	// --- optional, only consumed by walkBranchMergedBody (v2.5.0) ---
	/** Snapshot every tracked name's current state, for forking If arms. */
	snapshotState?(): Map<string, string>;
	/** Overwrite the live state from a snapshot, restoring it before the next arm. */
	restoreState?(snapshot: ReadonlyMap<string, string>): void;
	/** Write one tracked name's merged post-block state. */
	setState?(lowerName: string, value: string): void;
	/** The rule's good/init labels so the branch merge stays rule-agnostic. */
	lattice?: { init: string; good: string; unknown: string };
}

/**
 * Walks the straight-line statements of a procedure body: plain statements run
 * the rule's transitions in order, while nested blocks (If/For/Do/...) are not
 * entered - every tracked name touched anywhere inside them is demoted to
 * 'unknown' instead of guessing which runtime path executes.
 */
export function walkStraightLineBody(
	body: readonly BodyNode[],
	isInactive: (node: BodyNode) => boolean,
	hooks: StraightLineDataflowHooks,
): void {
	for (const node of body) {
		if (isInactive(node)) {
			continue;
		}
		if (isLeafStatement(node)) {
			hooks.onStatement(node);
			continue;
		}
		hooks.onBlock?.(node);
		if ('body' in node && Array.isArray(node.body)) {
			for (const lower of collectNestedTouches(node.body, isInactive, hooks)) {
				hooks.demoteToUnknown(lower);
			}
		}
	}
}

/**
 * Like walkStraightLineBody, but intersects the per-branch state of an
 * If/ElseIf/Else block instead of blanket-demoting every name it touches. Each
 * arm is walked from the block's entry state; a tracked name advances to its
 * 'good' state after the `If` only when it reaches 'good' on EVERY arm AND a
 * syntactic `else` arm is present, otherwise it follows the conservative
 * demotion. Names a balanced `If` never touches keep their entry state (the
 * precision win). For/Do/While/With/Select stay conservative (the current
 * blanket demotion). Callers must supply snapshotState/restoreState/setState/
 * lattice; without them an `If` is treated conservatively.
 *
 * Only sound for procedures WITHOUT unstructured control flow (labels, GoTo,
 * On Error, Resume): callers gate on procedureHasUnstructuredFlow and fall back
 * to walkStraightLineBody when it holds.
 */
export function walkBranchMergedBody(
	body: readonly BodyNode[],
	isInactive: (node: BodyNode) => boolean,
	hooks: StraightLineDataflowHooks,
): void {
	for (const node of body) {
		if (isInactive(node)) {
			continue;
		}
		if (isLeafStatement(node)) {
			hooks.onStatement(node);
			continue;
		}
		hooks.onBlock?.(node);
		if (
			node.kind === 'IfBlock' &&
			hooks.snapshotState &&
			hooks.restoreState &&
			hooks.setState &&
			hooks.lattice
		) {
			mergeIfBlock(node, isInactive, hooks);
			continue;
		}
		if ('body' in node && Array.isArray(node.body)) {
			for (const lower of collectNestedTouches(node.body, isInactive, hooks)) {
				hooks.demoteToUnknown(lower);
			}
		}
	}
}

/** Intersects the per-arm state of one If block (see walkBranchMergedBody). */
function mergeIfBlock(
	ifBlock: IfBlockNode,
	isInactive: (node: BodyNode) => boolean,
	hooks: StraightLineDataflowHooks,
): void {
	const touched = collectNestedTouches(ifBlock.body, isInactive, hooks);
	const hasElse = ifBlock.branches.some((branch) => branch.branchKind === 'else');
	if (!hasElse) {
		// No else arm: the empty fall-through path keeps the entry state, so a name
		// can only remain 'good' after the block if it was already 'good'. Reproduce
		// the existing conservative behavior by demoting every touched name.
		for (const lower of touched) {
			hooks.demoteToUnknown(lower);
		}
		return;
	}
	const entry = hooks.snapshotState!();
	const armStates: Map<string, string>[] = [];
	for (const branch of ifBlock.branches) {
		hooks.restoreState!(entry);
		walkBranchMergedBody(branch.body, isInactive, hooks);
		armStates.push(hooks.snapshotState!());
	}
	hooks.restoreState!(entry);
	const { unknown } = hooks.lattice!;
	for (const lower of touched) {
		const fallback = entry.get(lower) ?? unknown;
		hooks.setState!(lower, joinBranchStates(armStates, lower, fallback, hooks.lattice!));
	}
}

/**
 * Meet-toward-unknown join over an If block's arms for one tracked name: 'good'
 * only when every arm ends 'good'; any unknown arm or any disagreement collapses
 * to 'unknown'. Each arm's state is read inline (falling back to the name's entry
 * state) so no intermediate per-name array is allocated.
 */
function joinBranchStates(
	armStates: readonly ReadonlyMap<string, string>[],
	lower: string,
	fallback: string,
	lattice: { init: string; good: string; unknown: string },
): string {
	const { init, good, unknown } = lattice;
	let allGood = true;
	let allInit = true;
	for (const arm of armStates) {
		const state = arm.get(lower) ?? fallback;
		if (state === unknown) {
			return unknown;
		}
		if (state !== good) {
			allGood = false;
		}
		if (state !== init) {
			allInit = false;
		}
	}
	return allGood ? good : allInit ? init : unknown;
}

/** Recursively collects tracked names touched anywhere inside nested bodies. */
function collectNestedTouches(
	body: readonly BodyNode[],
	isInactive: (node: BodyNode) => boolean,
	hooks: Pick<StraightLineDataflowHooks, 'touchesInStatement'>,
): Set<string> {
	const out = new Set<string>();
	for (const node of body) {
		if (isInactive(node)) {
			continue;
		}
		if (isLeafStatement(node)) {
			for (const lower of hooks.touchesInStatement(node)) {
				out.add(lower);
			}
			continue;
		}
		if ('body' in node && Array.isArray(node.body)) {
			for (const lower of collectNestedTouches(node.body, isInactive, hooks)) {
				out.add(lower);
			}
		}
	}
	return out;
}

/**
 * Lowercased tracked locals passed as bare arguments of a call statement
 * (`Helper x`, `Call Helper(x)`), where ByRef passing may rebind them. `toks`
 * are the statement's significant tokens after any leading label.
 */
export function trackedLocalsPassedAsCallArguments(
	toks: readonly VbaToken[],
	isTracked: (lowerName: string) => boolean,
): Set<string> {
	if (toks.length < 2 || hasTopLevelAssignment(toks)) {
		return new Set();
	}
	const start = tokenWord(toks[0]) === 'call' ? 1 : 0;
	if (!tokenName(toks[start]) || toks[start + 1]?.rawText === '.') {
		return new Set();
	}
	const out = new Set<string>();
	for (let i = start + 1; i < toks.length; i++) {
		if (toks[i - 1]?.rawText === '.' || toks[i + 1]?.rawText === '.') {
			continue;
		}
		const lower = tokenName(toks[i])?.toLowerCase();
		if (lower && isTracked(lower)) {
			out.add(lower);
		}
	}
	return out;
}

/** True when a top-level '=' makes the statement an assignment, not a call. */
function hasTopLevelAssignment(toks: readonly VbaToken[]): boolean {
	let depth = 0;
	for (const tok of toks) {
		const raw = tok.rawText;
		if (raw === '(' || raw === '[') {
			depth++;
		} else if (raw === ')' || raw === ']') {
			depth--;
		} else if (depth === 0 && tok.kind === 'operator' && raw === '=') {
			return true;
		}
	}
	return false;
}
