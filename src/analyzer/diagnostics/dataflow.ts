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
import { blockHeaderLeaves, isLoopBlock, selectArms } from './blockHeaders';

const NO_NAMES: ReadonlySet<string> = new Set();

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
	source: string,
	body: readonly BodyNode[],
	isInactive: (node: BodyNode) => boolean,
	hooks: StraightLineDataflowHooks,
): void {
	walkBody(source, body, isInactive, hooks, false);
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
	source: string,
	body: readonly BodyNode[],
	isInactive: (node: BodyNode) => boolean,
	hooks: StraightLineDataflowHooks,
): void {
	walkBody(source, body, isInactive, hooks, true);
}

/** The walk both entry points share; merging If arms is the one place they differ. */
function walkBody(
	source: string,
	body: readonly BodyNode[],
	isInactive: (node: BodyNode) => boolean,
	hooks: StraightLineDataflowHooks,
	mergeIfBlocks: boolean,
	/** Names an enclosing loop changes: a nested block may run on a later pass. */
	loopTouched: ReadonlySet<string> = NO_NAMES,
): void {
	for (let i = 0; i < body.length; i++) {
		const node = body[i];
		if (isInactive(node)) {
			continue;
		}
		// A single-line If runs its statements on some passes only.
		if (mergeIfBlocks && isConditionalLeaf(node)) {
			for (const lower of loopTouched) {
				hooks.demoteToUnknown(lower);
			}
		}
		if (isSingleLineIfTail(node)) {
			const tail: LeafStatementNode[] = [];
			for (; i < body.length && isSingleLineIfTail(body[i]); i++) {
				tail.push(body[i] as LeafStatementNode);
			}
			i--;
			walkSingleLineIfTail(tail, hooks);
			continue;
		}
		if (isLeafStatement(node)) {
			hooks.onStatement(node);
			continue;
		}
		hooks.onBlock?.(node);
		if (mergeIfBlocks) {
			for (const lower of loopTouched) {
				hooks.demoteToUnknown(lower);
			}
		}
		if (
			mergeIfBlocks &&
			node.kind === 'IfBlock' &&
			hooks.snapshotState &&
			hooks.restoreState &&
			hooks.setState &&
			hooks.lattice
		) {
			mergeIfBlock(source, node, isInactive, hooks, loopTouched);
			continue;
		}
		if ('body' in node && Array.isArray(node.body)) {
			const touched = blockTouches(source, node, isInactive, hooks);
			if (mergeIfBlocks && hooks.snapshotState && hooks.restoreState) {
				walkBlockFromEntry(source, node, touched, isInactive, hooks, loopTouched);
			}
			for (const lower of touched) {
				hooks.demoteToUnknown(lower);
			}
		}
	}
}

/**
 * Checks the statements of a For, Do, While, Select or With block with the
 * state the block is entered with (issue #237), once its own lines have run:
 * a block that never touches a name leaves what is known about it as it was.
 * A statement directly in a loop's body runs on the first pass as it stands;
 * a block nested in the loop may run on a later pass, and forgets what the
 * loop changes before it is walked. Each Case of a Select starts from the
 * entry state. The state after the block is the caller's to set.
 */
function walkBlockFromEntry(
	source: string,
	node: BodyNode & { body: BodyNode[] },
	touched: ReadonlySet<string>,
	isInactive: (node: BodyNode) => boolean,
	hooks: StraightLineDataflowHooks,
	loopTouched: ReadonlySet<string>,
): void {
	for (const lower of headerTouches(source, node, hooks)) {
		hooks.demoteToUnknown(lower);
	}
	const entry = hooks.snapshotState!();
	if (node.kind === 'SelectBlock') {
		for (const arm of selectArms(source, node.body)) {
			hooks.restoreState!(entry);
			walkBody(source, arm, isInactive, hooks, true, loopTouched);
		}
	} else {
		walkBody(source, node.body, isInactive, hooks, true, isLoopBlock(node) ? new Set([...loopTouched, ...touched]) : loopTouched);
	}
	hooks.restoreState!(entry);
}

/** A single-line If, or a statement it runs after a colon. */
function isConditionalLeaf(node: BodyNode): boolean {
	return isLeafStatement(node) && (node.singleLineIfTail === true || (node.kind === 'Statement' && node.singleLineIfBranches !== undefined));
}

function isSingleLineIfTail(node: BodyNode): node is LeafStatementNode {
	return isLeafStatement(node) && node.singleLineIfTail === true;
}

/**
 * The statements a single-line If runs after a colon, `b` in `If x Then a: b`,
 * run only with its branch (MS-VBAL 5.4.2.9). They are checked on that path,
 * and afterwards the state is what a block If without Else leaves: as it was
 * before them, with every name they touch made unknown.
 */
function walkSingleLineIfTail(
	tail: readonly LeafStatementNode[],
	hooks: StraightLineDataflowHooks,
): void {
	const entry = hooks.snapshotState?.();
	if (entry && hooks.restoreState) {
		for (const stmt of tail) {
			hooks.onStatement(stmt);
		}
		hooks.restoreState(entry);
	}
	for (const stmt of tail) {
		for (const lower of hooks.touchesInStatement(stmt)) {
			hooks.demoteToUnknown(lower);
		}
	}
}

/** Intersects the per-arm state of one If block (see walkBranchMergedBody). */
function mergeIfBlock(
	source: string,
	ifBlock: IfBlockNode,
	isInactive: (node: BodyNode) => boolean,
	hooks: StraightLineDataflowHooks,
	loopTouched: ReadonlySet<string>,
): void {
	const touched = blockTouches(source, ifBlock, isInactive, hooks);
	const hasElse = ifBlock.branches.some((branch) => branch.branchKind === 'else');
	// Each arm is checked from the block's entry state, once its conditions
	// have run: `If TryGet(k, obj) Then` sets obj for the arm (issue #237).
	for (const lower of headerTouches(source, ifBlock, hooks)) {
		hooks.demoteToUnknown(lower);
	}
	const entry = hooks.snapshotState!();
	const armStates: Map<string, string>[] = [];
	for (const branch of ifBlock.branches) {
		hooks.restoreState!(entry);
		walkBody(source, branch.body, isInactive, hooks, true, loopTouched);
		armStates.push(hooks.snapshotState!());
	}
	hooks.restoreState!(entry);
	if (!hasElse) {
		// No else arm: the empty fall-through path keeps the entry state, so a name
		// can only remain 'good' after the block if it was already 'good'. Reproduce
		// the existing conservative behavior by demoting every touched name.
		for (const lower of touched) {
			hooks.demoteToUnknown(lower);
		}
		return;
	}
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
	source: string,
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
			for (const lower of blockTouches(source, node, isInactive, hooks)) {
				out.add(lower);
			}
		}
	}
	return out;
}

/** The tracked names a block's own lines pass on: `If TryGet(k, obj) Then`. */
function headerTouches(
	source: string,
	node: BodyNode,
	hooks: Pick<StraightLineDataflowHooks, 'touchesInStatement'>,
): Set<string> {
	const out = new Set<string>();
	for (const header of blockHeaderLeaves(source, node)) {
		for (const lower of hooks.touchesInStatement(header)) {
			out.add(lower);
		}
	}
	return out;
}

/**
 * The tracked names a block may change: those its body touches, and those
 * its own lines pass on, `If TryGet(k, obj) Then` and a For Each's control
 * variable among them (issue #237).
 */
function blockTouches(
	source: string,
	node: BodyNode,
	isInactive: (node: BodyNode) => boolean,
	hooks: Pick<StraightLineDataflowHooks, 'touchesInStatement'>,
): Set<string> {
	const out = collectNestedTouches(source, (node as { body: BodyNode[] }).body, isInactive, hooks);
	for (const header of blockHeaderLeaves(source, node)) {
		for (const lower of hooks.touchesInStatement(header)) {
			out.add(lower);
		}
	}
	return out;
}

/** The state a rule's own walk keeps, and how a block is allowed to change it. */
export interface BlockEnteringState<S> {
	/** A copy of the state, to enter each block and each If arm from. */
	snapshot(): S;
	/** Puts back a copy taken by snapshot. */
	restore(state: S): void;
	/** Drops what is known about these tracked names. */
	forget(names: ReadonlySet<string>): void;
	/** The tracked names a statement mentions, which a block may change. */
	touches(stmt: LeafStatementNode): Iterable<string>;
}

/**
 * A rule's own statement walk, entering blocks (issue #237). A statement
 * inside a block is visited with the state the block is entered with, once
 * the block's own lines have run: each If arm and each Case from that state,
 * a With's body once, and a loop's body as its first pass runs it. A block
 * or a single-line If nested in a loop may run on a later pass, so it first
 * forgets every name the loop changes (issue #238). After a block the state
 * is its entry state less every name it touches, so a block that never names
 * a variable keeps what is known about it.
 */
export function walkEnteringBlocks<S>(
	source: string,
	body: readonly BodyNode[],
	isInactive: (node: BodyNode) => boolean,
	visit: (node: BodyNode) => void,
	state: BlockEnteringState<S>,
	loopTouched: ReadonlySet<string> = NO_NAMES,
): void {
	for (const node of body) {
		if (isInactive(node)) {
			continue;
		}
		if (!('body' in node) || !Array.isArray(node.body)) {
			// A single-line If runs its statements on some passes only.
			if (isConditionalLeaf(node)) {
				state.forget(loopTouched);
			}
			visit(node);
			continue;
		}
		const hooks = { touchesInStatement: (stmt: LeafStatementNode) => state.touches(stmt) };
		const touched = blockTouches(source, node, isInactive, hooks);
		// An enclosing loop may have changed these on an earlier pass, and the
		// block's own lines run before its body.
		state.forget(loopTouched);
		state.forget(headerTouches(source, node, hooks));
		const entry = state.snapshot();
		if (node.kind === 'IfBlock') {
			for (const branch of node.branches) {
				state.restore(entry);
				walkEnteringBlocks(source, branch.body, isInactive, visit, state, loopTouched);
			}
		} else if (node.kind === 'SelectBlock') {
			for (const arm of selectArms(source, node.body)) {
				state.restore(entry);
				walkEnteringBlocks(source, arm, isInactive, visit, state, loopTouched);
			}
		} else {
			walkEnteringBlocks(source, node.body, isInactive, visit, state, isLoopBlock(node) ? new Set([...loopTouched, ...touched]) : loopTouched);
		}
		state.restore(entry);
		state.forget(touched);
	}
}

/**
 * Lowercased tracked locals passed as bare arguments of a call statement
 * (`Helper x`, `Call Helper(x)`), where ByRef passing may rebind them. `toks`
 * are the statement's significant tokens after any leading label.
 */
export function trackedLocalsNamedWhole(
	toks: readonly VbaToken[],
	spanStart: number,
	isTracked: (lowerName: string) => boolean,
	readOnlyIntrinsics: ReadonlySet<string>,
): Map<string, number> {
	const out = new Map<string, number>();
	if (toks.length < 2) {
		return out;
	}
	// A bare mention at the top level is an argument only in a call statement:
	// `Foo x`, `Call Foo(x)`, `obj.Method x`, or `.Method x` inside With. In
	// `Set a = b`, `Dim a As T`, or `If a Is Nothing` it is not.
	const head = tokenWord(toks[0]) === 'call' ? 1 : 0;
	const isCallStatement =
		(tokenName(toks[head]) !== undefined || toks[head]?.rawText === '.') &&
		!hasTopLevelAssignment(toks);
	let depth = 0;
	for (let i = 1; i < toks.length; i++) {
		const raw = toks[i].rawText;
		if (raw === '(' || raw === '[') {
			depth++;
			continue;
		}
		if (raw === ')' || raw === ']') {
			depth--;
			continue;
		}
		const lower = tokenName(toks[i])?.toLowerCase();
		if (!lower || !isTracked(lower) || out.has(lower)) {
			continue;
		}
		if (depth === 0 && !isCallStatement) {
			continue;
		}
		const prev = toks[i - 1]?.rawText;
		const next = toks[i + 1];
		if (prev === '.' || prev === '!' || next?.rawText === '(' || next?.rawText === '.' || next?.rawText === '!') {
			continue;
		}
		if (tokenWord(next) === 'is') {
			continue;
		}
		if (prev === '(' && readOnlyIntrinsics.has(tokenName(toks[i - 2])?.toLowerCase() ?? '')) {
			continue;
		}
		out.set(lower, spanStart + toks[i].start);
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
