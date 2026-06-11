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
import type { BodyNode, StatementNode } from '../parser/nodes';

/** Rule-specific hooks driving one straight-line dataflow walk. */
export interface StraightLineDataflowHooks {
	/** Applies one straight-line statement's transitions and diagnostics. */
	onStatement(stmt: StatementNode): void;
	/** Inspects one non-statement node before its body's touch demotion. */
	onBlock?(node: BodyNode): void;
	/** Lowercased tracked names one nested-block statement touches. */
	touchesInStatement(stmt: StatementNode): Iterable<string>;
	/** Demotes one tracked name to the rule's 'unknown' state. */
	demoteToUnknown(lowerName: string): void;
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
		if (node.kind === 'Statement') {
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
		if (node.kind === 'Statement') {
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
