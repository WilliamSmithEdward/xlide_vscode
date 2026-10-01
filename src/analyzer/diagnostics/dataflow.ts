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
import { statementTokensCached, tokenName, tokensWithoutLeadingLineNumber, tokenWord } from '../lexer/tokenHelpers';
import { statementLabelDeclarations, statementLabelReferences } from '../flow/procedureLabels';
import type { BodyNode, IfBlockNode, LeafStatementNode, Span } from '../parser/nodes';
import { isLeafStatement } from '../parser/nodes';
import { blockHeaderLeaves, isLoopBlock, selectArms } from './blockHeaders';
import { ifConditionTokens } from './conditionValue';

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
	/**
	 * Drops the rule's findings while true: the GoTo-following walk runs the
	 * body more than once to settle what each label is entered with, and
	 * reports on its last run only (issue #271).
	 */
	setSilent?(silent: boolean): void;
	/**
	 * What an If condition comes to from the rule's own state, undefined
	 * when that is not certain: `c Is Nothing` with c never Set. An arm the
	 * condition rules out is not walked, and one that always leaves ends
	 * the path (issue #273).
	 */
	knownCondition?(condition: readonly VbaToken[]): boolean | undefined;
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
	if (!walkFollowingJumps(source, body, isInactive, hooks)) {
		walkBody(source, body, isInactive, hooks, false);
	}
}

/** The most runs the GoTo-following walk takes to settle its labels. */
const MAX_JUMP_PASSES = 6;

/**
 * Walks the top-level statements following GoTo (issue #271, measured in
 * Excel 16.0). A label is entered with the state that falls into it, if the
 * statement before it can, merged with the state at each unconditional
 * top-level `GoTo` to it: `GoTo Setup` ... `Setup: Set c = ...: GoTo Use`
 * reaches `Use:` with c set, and `GoTo Use` past the Set reaches it with c
 * still Nothing. A jump that may or may not run - a GoTo in a single-line If
 * or a block, On Error GoTo, GoSub, Resume to a label, On ... GoTo - enters
 * its label with nothing known. Code nothing reaches is not checked. The
 * body is run until the labels settle, then once more to report. A Resume
 * with no label, which returns into the body anywhere, keeps the plain walk:
 * false is returned then, and when the rule cannot snapshot its state.
 */
function walkFollowingJumps(
	source: string,
	body: readonly BodyNode[],
	isInactive: (node: BodyNode) => boolean,
	hooks: StraightLineDataflowHooks,
): boolean {
	const { snapshotState, restoreState, lattice, setSilent } = hooks;
	if (!snapshotState || !restoreState || !lattice || !setSilent) {
		return false;
	}
	const leaves: LeafStatementNode[] = [];
	collectLeaves(body, isInactive, leaves);
	if (!leaves.some((leaf) => statementLabelReferences(source, leaf.span).length > 0 || statementLabelDeclarations(source, leaf.span).length > 0)) {
		return false;
	}
	if (leaves.some((leaf) => {
		const toks = statementTokensAfterLabel(source, leaf);
		return toks.some((tok, k) => tokenWord(tok) === 'resume' && (k + 1 >= toks.length || tokenWord(toks[k + 1]) === 'next' || toks[k + 1].kind === 'comment'));
	})) {
		return false;
	}
	const initial = snapshotState();
	const unknownState = new Map([...initial.keys()].map((key) => [key, lattice.unknown]));
	const merge = (states: readonly ReadonlyMap<string, string>[]): Map<string, string> => {
		const out = new Map<string, string>();
		for (const key of initial.keys()) {
			const values = new Set(states.map((state) => state.get(key) ?? lattice.unknown));
			out.set(key, values.size === 1 ? [...values][0] : lattice.unknown);
		}
		return out;
	};
	const run = (entries: ReadonlyMap<string, ReadonlyMap<string, string>>): Map<string, Map<string, string>> => {
		restoreState(initial);
		const reaching = new Map<string, Map<string, string>[]>();
		const arrive = (key: string, state: ReadonlyMap<string, string>): void => {
			reaching.set(key, [...(reaching.get(key) ?? []), new Map(state)]);
		};
		/** Walks one statement list; returns whether its end is reached. */
		const runList = (list: readonly BodyNode[], reachableIn: boolean): boolean => {
			let reachable = reachableIn;
			for (let i = 0; i < list.length; i++) {
				const node = list[i];
				if (isInactive(node)) {
					continue;
				}
				if (isLeafStatement(node)) {
					const labels = statementLabelDeclarations(source, node.span).map((label) => label.key);
					if (labels.length > 0) {
						const states: ReadonlyMap<string, string>[] = [...(reachable ? [snapshotState()] : []), ...labels.flatMap((key) => (entries.get(key) ? [entries.get(key)!] : []))];
						reachable = states.length > 0;
						if (reachable) {
							restoreState(merge(states));
						}
					}
				}
				if (!reachable) {
					continue;
				}
				const guard = knownSingleLineIf(source, list, i, hooks);
				if (guard?.known === false) {
					i += guard.group.length - 1;
					continue;
				}
				if (guard?.leaves) {
					// It runs, and a GoTo in it arrives with what holds after it.
					i += guard.group.length - 1;
					hooks.onStatement(guard.group[0]);
					walkSingleLineIfTail(guard.group.slice(1), hooks);
					for (const stmt of guard.group) {
						for (const ref of statementLabelReferences(source, stmt.span)) {
							arrive(ref.key, ref.statementKind === 'goto' ? snapshotState() : unknownState);
						}
					}
					reachable = false;
					continue;
				}
				if (isSingleLineIfTail(node)) {
					const tail: LeafStatementNode[] = [];
					for (; i < list.length && isSingleLineIfTail(list[i]); i++) {
						tail.push(list[i] as LeafStatementNode);
					}
					i--;
					walkSingleLineIfTail(tail, hooks);
					for (const stmt of tail) {
						for (const ref of statementLabelReferences(source, stmt.span)) {
							arrive(ref.key, unknownState);
						}
					}
					continue;
				}
				if (isLeafStatement(list[i])) {
					// isSingleLineIfTail's guard narrows node away; it is a leaf here.
					const leaf = list[i] as LeafStatementNode;
					hooks.onStatement(leaf);
					const toks = statementTokensAfterLabel(source, leaf);
					const head = tokenWord(toks[0]);
					// A single-line If starts with If, so only a plain GoTo, Exit,
					// Resume, Return, End or Err.Raise ends the path here.
					for (const ref of statementLabelReferences(source, leaf.span)) {
						arrive(ref.key, ref.statementKind === 'goto' && head === 'goto' ? snapshotState() : unknownState);
					}
					if (leavesTheList(source, leaf.span)) {
						reachable = false;
					}
					continue;
				}
				hooks.onBlock?.(node);
				if (!('body' in node) || !Array.isArray(node.body)) {
					continue;
				}
				// An If runs one arm, or none without an Else; a Select one Case,
				// or none without Case Else; a With its body once. Each starts from
				// the block's entry state, and what follows merges where they end.
				// A loop may run any number of times, so it forgets what it touches.
				let arms: BodyNode[][] | undefined = node.kind === 'IfBlock' ? node.branches.map((branch) => branch.body)
					: node.kind === 'SelectBlock' ? selectArms(source, node.body)
						: node.kind === 'WithBlock' ? [node.body] : undefined;
				if (!arms) {
					// A GoTo in a loop leaves with the state the loop started
					// with, less whatever the loop may have changed.
					const touched = blockTouches(source, node, isInactive, hooks);
					const leaving = snapshotState();
					for (const lower of touched) {
						leaving.set(lower, lattice.unknown);
					}
					const nested: LeafStatementNode[] = [];
					collectLeaves(node.body as BodyNode[], isInactive, nested);
					for (const leaf of nested) {
						for (const ref of statementLabelReferences(source, leaf.span)) {
							arrive(ref.key, leaving);
						}
					}
					for (const lower of touched) {
						hooks.demoteToUnknown(lower);
					}
					continue;
				}
				for (const lower of headerTouches(source, node, hooks)) {
					hooks.demoteToUnknown(lower);
				}
				// An If runs only the arms its known conditions allow.
				const ifArms = node.kind === 'IfBlock' ? ifArmsThatMayRun(source, node, hooks) : undefined;
				if (ifArms) {
					arms = ifArms.arms.map((branch) => branch.body);
				}
				const entry = snapshotState();
				const ends: Map<string, string>[] = [];
				for (const arm of arms) {
					restoreState(entry);
					if (runList(arm, true)) {
						ends.push(snapshotState());
					}
				}
				const exhaustive = node.kind === 'WithBlock'
					|| ifArms?.exhaustive === true
					|| (node.kind === 'SelectBlock' && node.body.some((stmt) => isLeafStatement(stmt) && /^\s*case\s+else\b/i.test(source.slice(stmt.span.start, stmt.span.end))));
				if (!exhaustive) {
					ends.push(entry);
				}
				reachable = ends.length > 0;
				if (reachable) {
					restoreState(merge(ends));
				}
			}
			return reachable;
		};
		runList(body, true);
		return new Map([...reaching].map(([key, states]) => [key, merge(states)]));
	};
	let entries = new Map<string, Map<string, string>>();
	setSilent(true);
	for (let pass = 0; pass < MAX_JUMP_PASSES; pass++) {
		const next = run(entries);
		const settled = next.size === entries.size && [...next].every(([key, state]) => {
			const before = entries.get(key);
			return before !== undefined && [...state].every(([name, value]) => before.get(name) === value);
		});
		entries = next;
		if (settled) {
			break;
		}
	}
	setSilent(false);
	run(entries);
	return true;
}

function collectLeaves(body: readonly BodyNode[], isInactive: (node: BodyNode) => boolean, out: LeafStatementNode[]): void {
	for (const node of body) {
		if (isInactive(node)) {
			continue;
		}
		if (isLeafStatement(node)) {
			out.push(node);
		} else if ('body' in node && Array.isArray(node.body)) {
			collectLeaves(node.body as BodyNode[], isInactive, out);
		}
	}
}

/** A statement's tokens after any leading label or line number. */
function statementTokensAfterLabel(source: string, node: LeafStatementNode): VbaToken[] {
	return tokensAfterLabel(source, node.span);
}

function tokensAfterLabel(source: string, span: Span): VbaToken[] {
	let toks = tokensWithoutLeadingLineNumber(statementTokensCached(source, span)).filter((tok) => tok.kind !== 'comment');
	if (statementLabelDeclarations(source, span).length > 0 && toks[1]?.rawText === ':') {
		toks = toks.slice(2);
	}
	return toks;
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

/**
 * The walk both entry points share; merging If arms is the one place they
 * differ. Returns whether the end of the body is reached: with If arms
 * merged, a statement that always leaves ends the path (issue #273).
 */
function walkBody(
	source: string,
	body: readonly BodyNode[],
	isInactive: (node: BodyNode) => boolean,
	hooks: StraightLineDataflowHooks,
	mergeIfBlocks: boolean,
	/** Names an enclosing loop changes: a nested block may run on a later pass. */
	loopTouched: ReadonlySet<string> = NO_NAMES,
): boolean {
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
		const guard = mergeIfBlocks ? knownSingleLineIf(source, body, i, hooks) : undefined;
		if (guard?.known === false) {
			i += guard.group.length - 1;
			continue;
		}
		if (guard?.leaves) {
			hooks.onStatement(guard.group[0]);
			walkSingleLineIfTail(guard.group.slice(1), hooks);
			return false;
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
			// isSingleLineIfTail's guard narrows node away; it is a leaf here.
			if (mergeIfBlocks && leavesTheList(source, (body[i] as LeafStatementNode).span)) {
				return false;
			}
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
			if (!mergeIfBlock(source, node, isInactive, hooks, loopTouched)) {
				return false;
			}
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
	return true;
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

/** Statement heads after which the rest of the list does not run. */
const LIST_LEAVING_HEADS: ReadonlySet<string> = new Set(['exit', 'goto', 'resume', 'return']);

/**
 * Whether a statement always leaves the list it is in: Exit, GoTo,
 * Resume, Return, a bare End, and Err.Raise unless the procedure resumes
 * past errors (issue #273).
 */
export function leavesTheList(source: string, span: Span, raiseLeaves = true): boolean {
	const toks = tokensAfterLabel(source, span);
	const head = tokenWord(toks[0]);
	return LIST_LEAVING_HEADS.has(head) || (head === 'end' && toks.length === 1)
		|| (raiseLeaves && head === 'err' && toks[1]?.rawText === '.' && tokenWord(toks[2]) === 'raise');
}

/**
 * A single-line If with no Else and the statements after its colons, when
 * the rule knows its condition (issue #273): false runs none of them, and
 * true runs them all in order, which leaves when one of them does.
 */
function knownSingleLineIf(
	source: string,
	list: readonly BodyNode[],
	i: number,
	hooks: StraightLineDataflowHooks,
): { group: LeafStatementNode[]; known: boolean; leaves: boolean } | undefined {
	const node = list[i];
	if (!hooks.knownCondition || node.kind !== 'Statement' || node.singleLineIfBranches?.length !== 1) {
		return undefined;
	}
	const condition = ifConditionTokens(tokensAfterLabel(source, node.span));
	const known = condition ? hooks.knownCondition(condition) : undefined;
	if (known === undefined) {
		return undefined;
	}
	const group: LeafStatementNode[] = [node];
	for (let k = i + 1; k < list.length && isSingleLineIfTail(list[k]); k++) {
		group.push(list[k] as LeafStatementNode);
	}
	const branch = node.singleLineIfBranches[0];
	const leaves = known && group.some((stmt, k) => leavesTheList(source, k === 0 ? branch : stmt.span));
	return { group, known, leaves };
}

/**
 * The arms of an If block that may run, and whether one of them always
 * does (issue #273). A condition known false drops its arm, and one known
 * true, or an Else, drops every arm after it.
 */
function ifArmsThatMayRun(
	source: string,
	ifBlock: IfBlockNode,
	hooks: StraightLineDataflowHooks,
): { arms: IfBlockNode['branches']; exhaustive: boolean } {
	const arms: IfBlockNode['branches'] = [];
	for (const branch of ifBlock.branches) {
		if (branch.branchKind === 'else') {
			arms.push(branch);
			return { arms, exhaustive: true };
		}
		const condition = hooks.knownCondition ? ifConditionTokens(tokensAfterLabel(source, branch.headerSpan)) : undefined;
		const known = condition ? hooks.knownCondition!(condition) : undefined;
		if (known !== false) {
			arms.push(branch);
		}
		if (known === true) {
			return { arms, exhaustive: true };
		}
	}
	return { arms, exhaustive: false };
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

/**
 * Intersects the per-arm state of one If block (see walkBranchMergedBody).
 * An arm that always leaves takes no part, and false is returned when
 * no path goes past the block (issue #273).
 */
function mergeIfBlock(
	source: string,
	ifBlock: IfBlockNode,
	isInactive: (node: BodyNode) => boolean,
	hooks: StraightLineDataflowHooks,
	loopTouched: ReadonlySet<string>,
): boolean {
	const touched = blockTouches(source, ifBlock, isInactive, hooks);
	// Each arm is checked from the block's entry state, once its conditions
	// have run: `If TryGet(k, obj) Then` sets obj for the arm (issue #237).
	for (const lower of headerTouches(source, ifBlock, hooks)) {
		hooks.demoteToUnknown(lower);
	}
	// Only the arms its known conditions allow run; when that is one arm
	// that always runs, the state after the block is the state it ends with.
	const { arms, exhaustive } = ifArmsThatMayRun(source, ifBlock, hooks);
	if (arms.length === 0) {
		return true;
	}
	if (arms.length === 1 && exhaustive) {
		return walkBody(source, arms[0].body, isInactive, hooks, true, loopTouched);
	}
	const entry = hooks.snapshotState!();
	const armStates: Map<string, string>[] = [];
	for (const branch of arms) {
		hooks.restoreState!(entry);
		if (walkBody(source, branch.body, isInactive, hooks, true, loopTouched)) {
			armStates.push(hooks.snapshotState!());
		}
	}
	hooks.restoreState!(entry);
	if (armStates.length === 0) {
		// Every arm leaves: only the path that skips the block goes on.
		return !exhaustive;
	}
	if (!exhaustive) {
		// No else arm: the empty fall-through path keeps the entry state, so a name
		// can only remain 'good' after the block if it was already 'good'. Reproduce
		// the existing conservative behavior by demoting every touched name.
		for (const lower of touched) {
			hooks.demoteToUnknown(lower);
		}
		return true;
	}
	const { unknown } = hooks.lattice!;
	for (const lower of touched) {
		const fallback = entry.get(lower) ?? unknown;
		hooks.setState!(lower, joinBranchStates(armStates, lower, fallback, hooks.lattice!));
	}
	return true;
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
	/** Called as a block is entered, before its own lines run: a For Each header reads its source here. */
	enter?(node: BodyNode): void;
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
		state.enter?.(node);
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
