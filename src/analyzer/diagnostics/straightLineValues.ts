// The assignment each statement of a procedure sees for a local (issue #180).
//
// The value rules (division-by-zero, string-arithmetic-coercion,
// variant-value-misuse, array-subscript-out-of-bounds, runtime-argument-value)
// read a local's value from the procedure as a whole: a local every
// assignment gives the same literal. A second assignment anywhere turned them
// off, even one after the line that fails: `d = 0: x = 10 / d: d = 2`. This
// walk follows each statement list in order and records, for every leaf
// statement, the last `x = value` that reaches it with nothing between able
// to change x. The rules read that first and fall back to the procedure-wide
// value.
//
// What ends a value: another assignment (the new one replaces it), passing
// the name whole to a call (ByRef), a statement that writes it another way
// (Set, ReDim, Erase, Input #, Get #, Line Input #, LSet, RSet, Mid =), and
// any block that touches the name, since a loop or a branch may or may not
// have run it. A label ends every value, because a GoTo may arrive there from
// anywhere, and so does a GoSub, which may run any statement of the
// procedure. A block keeps the values it never touches, inside and after it.

import type { VbaToken } from '../lexer/tokenKinds';
import type { ConditionalActivityTracker } from '../conditional/conditionalCompilation';
import type { BodyNode, IfBlockNode, LeafStatementNode, Span } from '../parser/nodes';
import { isLeafStatement } from '../parser/nodes';
import { statementLabelDeclaration } from '../flow/procedureLabels';
import { parseVbaIntegerLiteral } from '../constants/integerConstantExpression';
import { trackedLocalsNamedWhole } from './dataflow';
import { isLoopBlock } from './blockHeaders';
import {
	bareAssignmentTarget,
	blockFooterLineSpan,
	blockHeaderLineSpan,
	isInactiveNode,
	rawExpressionTokens,
	statementAndBranchSpans,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
} from './walker';

/** The value tokens of each local's reaching assignment, by lowercased name. */
export type ReachingAssignments = ReadonlyMap<string, readonly VbaToken[]>;

const NONE: ReachingAssignments = new Map();

/** Statement heads that write every name they mention. */
const WRITING_HEADS: ReadonlySet<string> = new Set(['set', 'redim', 'erase', 'input', 'get', 'line', 'lset', 'rset', 'mid', 'mid$']);

/** VBA functions that only read an argument named whole. */
const READ_ONLY_INTRINSICS: ReadonlySet<string> = new Set(['lbound', 'ubound', 'isarray', 'len', 'lenb', 'isempty', 'isnull', 'isnumeric', 'typename', 'vartype']);

/**
 * For each statement of the body that some straight-line assignment
 * reaches, the reaching assignments; for a block, those that reach its
 * header. A statement not in the map has none. `initial` is what holds as
 * the procedure starts: each local's declared default, so `x = 1 / x` reads
 * the 0 x starts with (issue #259).
 */
export function straightLineAssignments(
	source: string,
	body: readonly BodyNode[],
	activity: ConditionalActivityTracker | undefined,
	initial: ReachingAssignments = NONE,
): ReadonlyMap<BodyNode, ReachingAssignments> {
	// Six rules ask for the same procedure in one pass; a parse makes a new
	// body, so the body is the key, with what holds at the start.
	const key = [...initial].map(([name, value]) => `${name}=${value.map((tok) => tok.rawText).join(' ')}`).sort().join('\n');
	const byStart = WALKS.get(body) ?? new Map<string, CachedWalk>();
	WALKS.set(body, byStart);
	const cached = byStart.get(key);
	if (cached && cached.source === source && cached.activity === activity) {
		return cached.result;
	}
	const out = new Map<BodyNode, ReachingAssignments>();
	walkList(source, body, initial, activity, out);
	byStart.set(key, { source, activity, result: out });
	return out;
}

interface CachedWalk {
	source: string;
	activity: ConditionalActivityTracker | undefined;
	result: ReadonlyMap<BodyNode, ReachingAssignments>;
}

const WALKS = new WeakMap<readonly BodyNode[], Map<string, CachedWalk>>();

/**
 * Walks one statement list from `entry` and returns what holds after it. The
 * map is replaced, never changed in place, so every statement can keep the
 * one it saw.
 */
function walkList(
	source: string,
	list: readonly BodyNode[],
	entry: ReachingAssignments,
	activity: ConditionalActivityTracker | undefined,
	out: Map<BodyNode, ReachingAssignments>,
	caseResets = false,
): ReachingAssignments {
	let current = entry;
	for (let i = 0; i < list.length; i++) {
		const node = list[i];
		if (isInactiveNode(activity, node) || node.kind === 'VariableGroup' || node.kind === 'ConditionalDirective') {
			continue;
		}
		if (!isLeafStatement(node)) {
			// What holds as the block starts: a For reads its bounds here
			// (issue #200).
			record(out, node, current);
			current = walkBlock(source, node, current, activity, out);
			continue;
		}
		if (statementLabelDeclaration(source, node.span)) {
			current = NONE;
		}
		if (caseResets && tokenText(statementTokensAfterLeadingLabel(source, node.span)[0]) === 'case') {
			// A Select's arms are exclusive: each starts where the block did.
			current = entry;
		}
		if (node.kind === 'Statement' && node.singleLineIfBranches) {
			// A single-line If and the statements after its colons run only
			// with their branch. They see what held before the If, less
			// whatever the If itself changes, and so does what follows.
			const group: LeafStatementNode[] = [node];
			while (i + 1 < list.length && isLeafStatement(list[i + 1]) && (list[i + 1] as LeafStatementNode).singleLineIfTail) {
				group.push(list[++i] as LeafStatementNode);
			}
			const touched = touchedBy(source, group);
			current = touched === 'all' ? NONE : without(current, touched);
			for (const stmt of group) {
				record(out, stmt, current);
			}
			continue;
		}
		record(out, node, current);
		current = afterStatement(source, node.span, current);
	}
	return current;
}

function walkBlock(
	source: string,
	node: BodyNode,
	entry: ReachingAssignments,
	activity: ConditionalActivityTracker | undefined,
	out: Map<BodyNode, ReachingAssignments>,
): ReachingAssignments {
	if (!('body' in node) || !Array.isArray(node.body)) {
		return entry;
	}
	const touched = touchedInBlock(source, node, activity);
	const after = touched === 'all' ? NONE : without(entry, touched);
	// An If arm, a Case or a With body runs once, from the state the block is
	// entered with; a loop's body may run again with what it changed (issue
	// #259: `If True Then x = 1 / x` reads the 0 x starts with).
	const inside = isLoopBlock(node) || touched === 'all' ? after : entry;
	if (node.kind === 'IfBlock') {
		for (const branch of (node as IfBlockNode).branches) {
			walkList(source, branch.body, inside, activity, out);
		}
	} else {
		walkList(source, node.body as BodyNode[], inside, activity, out, node.kind === 'SelectBlock');
	}
	const final = touched === 'all' ? undefined : forCounterFinalValue(source, node, activity);
	if (final !== undefined) {
		const next = new Map(after);
		next.set(final.name, rawExpressionTokens(String(final.value)));
		return next;
	}
	return after;
}

/**
 * What a For counter holds after a loop of literal bounds runs to its end
 * (issue #263, measured in Excel 16.0): one step past its last pass, or the
 * start when no pass runs. `For i = 0 To 3 ... Next` leaves i at 4, so
 * `a(i)` on a Dim a(3) raises 9. A body that may leave the loop, or that
 * writes the counter, keeps it unknown.
 */
function forCounterFinalValue(
	source: string,
	node: BodyNode,
	activity: ConditionalActivityTracker | undefined,
): { name: string; value: number } | undefined {
	if (node.kind !== 'ForBlock' || node.each || !node.controlVariable) {
		return undefined;
	}
	const toks = statementTokensAfterLeadingLabel(source, blockHeaderLineSpan(source, node.span)).filter((tok) => tok.kind !== 'comment');
	const eq = toks.findIndex((tok) => tok.rawText === '=');
	const to = toks.findIndex((tok) => tokenText(tok) === 'to');
	const stepAt = toks.findIndex((tok) => tokenText(tok) === 'step');
	const start = eq > 0 && to > eq ? signedInteger(toks.slice(eq + 1, to)) : undefined;
	const limit = to > 0 ? signedInteger(toks.slice(to + 1, stepAt > 0 ? stepAt : toks.length)) : undefined;
	const step = stepAt > 0 ? signedInteger(toks.slice(stepAt + 1)) : 1;
	if (start === undefined || limit === undefined || step === undefined || step === 0) {
		return undefined;
	}
	const lower = node.controlVariable.toLowerCase();
	if (loopBodyMayLeaveOrWrite(source, node.body as BodyNode[], lower, activity)) {
		return undefined;
	}
	const passes = step > 0 ? (start <= limit ? Math.floor((limit - start) / step) + 1 : 0) : (start >= limit ? Math.floor((start - limit) / -step) + 1 : 0);
	return { name: lower, value: start + passes * step };
}

function signedInteger(toks: readonly VbaToken[]): number | undefined {
	const negative = toks.length === 2 && toks[0].rawText === '-';
	const literal = toks[negative ? 1 : 0];
	if (toks.length !== (negative ? 2 : 1) || literal.kind !== 'integerLiteral') {
		return undefined;
	}
	const value = parseVbaIntegerLiteral(literal.rawText);
	return value === undefined ? undefined : negative ? -value : value;
}

/** Statement heads that may leave a loop or jump within the procedure. */
const LEAVING_HEADS: ReadonlySet<string> = new Set(['exit', 'goto', 'gosub', 'resume', 'return', 'on']);

/** Whether a loop body may leave it early, jump, or write the counter. */
function loopBodyMayLeaveOrWrite(source: string, body: readonly BodyNode[], lower: string, activity: ConditionalActivityTracker | undefined): boolean {
	for (const node of body) {
		if (isInactiveNode(activity, node)) {
			continue;
		}
		if (!isLeafStatement(node)) {
			if (node.kind === 'ForBlock' && node.controlVariable?.toLowerCase() === lower) {
				return true;
			}
			const header = statementTokensAfterLeadingLabel(source, blockHeaderLineSpan(source, node.span));
			if ([...passedWhole(header, node.span.start)].includes(lower)) {
				return true;
			}
			if ('body' in node && Array.isArray(node.body) && loopBodyMayLeaveOrWrite(source, node.body as BodyNode[], lower, activity)) {
				return true;
			}
			continue;
		}
		if (statementLabelDeclaration(source, node.span)) {
			return true;
		}
		for (const span of statementAndBranchSpans(node)) {
			const toks = statementTokensAfterLeadingLabel(source, span);
			const head = tokenText(toks[0]);
			if (LEAVING_HEADS.has(head) || (head === 'end' && toks.length === 1)) {
				return true;
			}
			if (WRITING_HEADS.has(head) && mentionedNames(toks).has(lower)) {
				return true;
			}
			if (bareAssignmentTarget(source, span)?.name.toLowerCase() === lower || [...passedWhole(toks, span.start)].includes(lower)) {
				return true;
			}
		}
	}
	return false;
}

/** What holds after one plain statement runs. */
function afterStatement(source: string, span: Span, before: ReachingAssignments): ReachingAssignments {
	const toks = statementTokensAfterLeadingLabel(source, span);
	const head = tokenText(toks[0]);
	if (head === 'gosub') {
		return NONE;
	}
	if (WRITING_HEADS.has(head) && !(head === 'line' && tokenText(toks[1]) !== 'input')) {
		return without(before, mentionedNames(toks));
	}
	let after = without(before, passedWhole(toks, span.start));
	const bare = bareAssignmentTarget(source, span);
	if (bare) {
		const next = new Map(after);
		next.set(bare.name.toLowerCase(), bare.valueTokens.filter((tok) => tok.kind !== 'comment'));
		after = next;
	}
	return after;
}

/** Every name a block may change, header and footer lines included, or 'all'. */
function touchedInBlock(
	source: string,
	block: BodyNode,
	activity: ConditionalActivityTracker | undefined,
): Set<string> | 'all' {
	const names = new Set<string>();
	if (block.kind === 'ForBlock' && block.controlVariable) {
		names.add(block.controlVariable.toLowerCase());
	}
	for (const span of [blockHeaderLineSpan(source, block.span), blockFooterLineSpan(source, block.span)]) {
		for (const lower of passedWhole(statementTokensAfterLeadingLabel(source, span), span.start)) {
			names.add(lower);
		}
	}
	const visit = (list: readonly BodyNode[]): boolean => {
		for (const node of list) {
			if (isInactiveNode(activity, node)) {
				continue;
			}
			if (isLeafStatement(node)) {
				// A label inside the block can be reached from anywhere, and
				// what follows the block then runs with whatever that path held.
				if (statementLabelDeclaration(source, node.span)) {
					return true;
				}
				const touched = touchedBy(source, [node]);
				if (touched === 'all') {
					return true;
				}
				for (const lower of touched) {
					names.add(lower);
				}
				continue;
			}
			if (node.kind === 'ForBlock' && node.controlVariable) {
				names.add(node.controlVariable.toLowerCase());
			}
			if ('body' in node && Array.isArray(node.body)) {
				if (node.kind === 'IfBlock') {
					for (const branch of node.branches) {
						for (const lower of passedWhole(statementTokensAfterLeadingLabel(source, branch.headerSpan), branch.headerSpan.start)) {
							names.add(lower);
						}
					}
				}
				if (visit(node.body as BodyNode[])) {
					return true;
				}
			}
		}
		return false;
	};
	return 'body' in block && visit(block.body as BodyNode[]) ? 'all' : names;
}

/** Every name the statements may change, or 'all' for a GoSub. */
function touchedBy(source: string, stmts: readonly LeafStatementNode[]): Set<string> | 'all' {
	const names = new Set<string>();
	for (const stmt of stmts) {
		for (const span of statementAndBranchSpans(stmt)) {
			const toks = statementTokensAfterLeadingLabel(source, span);
			const head = tokenText(toks[0]);
			if (head === 'gosub') {
				return 'all';
			}
			const changed = WRITING_HEADS.has(head) ? mentionedNames(toks) : passedWhole(toks, span.start);
			for (const lower of changed) {
				names.add(lower);
			}
			const bare = bareAssignmentTarget(source, span);
			if (bare) {
				names.add(bare.name.toLowerCase());
			}
		}
	}
	return names;
}

function passedWhole(toks: readonly VbaToken[], spanStart: number): Iterable<string> {
	return trackedLocalsNamedWhole(toks, spanStart, () => true, READ_ONLY_INTRINSICS).keys();
}

function mentionedNames(toks: readonly VbaToken[]): Set<string> {
	const names = new Set<string>();
	for (const tok of toks) {
		const lower = tokenName(tok)?.toLowerCase();
		if (lower) {
			names.add(lower);
		}
	}
	return names;
}

function without(map: ReachingAssignments, names: Iterable<string>): ReachingAssignments {
	let next: Map<string, readonly VbaToken[]> | undefined;
	for (const lower of names) {
		if ((next ?? map).has(lower)) {
			next ??= new Map(map);
			next.delete(lower);
		}
	}
	return next ?? map;
}

function record(out: Map<BodyNode, ReachingAssignments>, stmt: BodyNode, current: ReachingAssignments): void {
	if (current.size > 0) {
		out.set(stmt, current);
	}
}
