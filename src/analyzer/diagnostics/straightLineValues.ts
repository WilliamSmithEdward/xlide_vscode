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
import { jumpTargetLabelDeclaration, statementLabelDeclaration, statementLabelReferences } from '../flow/procedureLabels';
import { parseVbaIntegerLiteral } from '../constants/integerConstantExpression';
import { leavesTheList, trackedLocalsNamedWhole } from './dataflow';
import { isLoopBlock, selectArms } from './blockHeaders';
import { conditionValue, ifConditionTokens, type ConditionFacts } from './conditionValue';
import { matchParenFrom, splitTopLevelTokenGroups } from '../lexer/tokenHelpers';
import { resolveRuntimeFunction } from '../runtime/vbaRuntime';
import { calleeKeepsArgument } from './calleeArguments';
import {
	bareAssignmentTarget,
	blockFooterLineSpan,
	blockHeaderLineSpan,
	isInactiveNode,
	rawExpressionTokens,
	statementAndBranchSpans,
	statementTokens,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
} from './walker';

/**
 * The value tokens of each local's reaching assignment, by lowercased name,
 * and of a local array's element by `elementKey`: `a(0) = Null` is under
 * "a(0)".
 */
export type ReachingAssignments = ReadonlyMap<string, readonly VbaToken[]>;

const NONE: ReachingAssignments = new Map();

/** The key a local array's element is held under: "a(0)". */
export function elementKey(lower: string, index: number): string {
	return `${lower}(${index})`;
}

/**
 * What the walk knows an object local holds, by identity: Nothing, from
 * `Set c = Nothing` or a local never set, and a Collection nothing has
 * added to yet, from `Set c = New Collection` or `Dim c As New Collection`.
 * A loop over the one, or until the other is Nothing, runs no pass (issue
 * #483). Any other mention of the name ends it.
 */
export const OBJECT_NOTHING: readonly VbaToken[] = rawExpressionTokens('Nothing');
export const EMPTY_COLLECTION: readonly VbaToken[] = rawExpressionTokens('New Collection');

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
	return cachedWalk(source, body, activity, initial).result;
}

/**
 * The statements of the body that never run, because a guard whose value the
 * walk knows decides against them (issue #273): `n = 0: If n > 0 Then ...`,
 * the arms after `Case 0` with the selector 0, and everything after
 * `If d = 0 Then Exit Function` with d still 0. A label ends it, since a
 * GoTo may arrive there. A block's own statements are in it when the whole
 * block never runs.
 */
/**
 * The spans of the one-line If branches that never run, because the walk
 * knows the condition (issue #430): the Else of `If x = 2 Then ... Else ...`
 * with x still 2. The If itself runs, so it is not in the unreachable set.
 */
export function straightLineDeadBranches(
	source: string,
	body: readonly BodyNode[],
	activity: ConditionalActivityTracker | undefined,
	initial: ReachingAssignments = NONE,
): readonly Span[] {
	return cachedWalk(source, body, activity, initial).deadSpans;
}

/**
 * What holds as the body runs off its end, from `initial`, with the
 * statements and one-line If branches the walk found never run. The end
 * state is undefined when no path reaches it (issue #562).
 */
export function straightLineExit(
	source: string,
	body: readonly BodyNode[],
	activity: ConditionalActivityTracker | undefined,
	initial: ReachingAssignments,
): { exit: ReachingAssignments | undefined; dead: ReadonlySet<BodyNode>; deadSpans: readonly Span[] } {
	const walk = cachedWalk(source, body, activity, initial);
	return { exit: walk.exit, dead: walk.dead, deadSpans: walk.deadSpans };
}

export function straightLineUnreachable(
	source: string,
	body: readonly BodyNode[],
	activity: ConditionalActivityTracker | undefined,
	initial: ReachingAssignments = NONE,
): ReadonlySet<BodyNode> {
	return cachedWalk(source, body, activity, initial).dead;
}

function cachedWalk(
	source: string,
	body: readonly BodyNode[],
	activity: ConditionalActivityTracker | undefined,
	initial: ReachingAssignments,
): CachedWalk {
	// Six rules ask for the same procedure in one pass; a parse makes a new
	// body, so the body is the key, with what holds at the start.
	let key = START_KEYS.get(initial);
	if (key === undefined) {
		key = [...initial].map(([name, value]) => `${name}=${value.map((tok) => tok.rawText).join(' ')}`).sort().join('\n');
		START_KEYS.set(initial, key);
	}
	const byStart = WALKS.get(body) ?? new Map<string, CachedWalk>();
	WALKS.set(body, byStart);
	const cached = byStart.get(key);
	if (cached && cached.source === source && cached.activity === activity) {
		return cached;
	}
	const out = new Map<BodyNode, ReachingAssignments>();
	const dead = new Set<BodyNode>();
	// Under On Error Resume Next, Err.Raise goes on to the next line.
	const text = body.length > 0 ? source.slice(body[0].span.start, body[body.length - 1].span.end) : '';
	const deadSpans: Span[] = [];
	// The walk is synchronous, so its arrays can sit beside it for passedWhole.
	const outer = walkArrays;
	const outerProcedures = walkProcedures;
	const outerElements = walkElements;
	walkArrays = localArrayNames(body, activity);
	let exit: ReachingAssignments;
	walkProcedures = moduleProcedureNames(source);
	walkElements = /\)\s*=\s*null\b/i.test(text);
	try {
		exit = walkList(source, body, initial, activity, { out, dead, deadSpans, raiseLeaves: !/\bon\s+error\s+resume\s+next\b/i.test(text), referenced: referencedLabels(source, body, activity) });
	} finally {
		walkArrays = outer;
		walkProcedures = outerProcedures;
		walkElements = outerElements;
	}
	const walk: CachedWalk = { source, activity, result: out, dead, deadSpans, exit: exit === UNREACHED ? undefined : exit };
	byStart.set(key, walk);
	return walk;
}

interface CachedWalk {
	source: string;
	activity: ConditionalActivityTracker | undefined;
	result: ReadonlyMap<BodyNode, ReachingAssignments>;
	dead: ReadonlySet<BodyNode>;
	deadSpans: readonly Span[];
	/** What holds as the body runs off its end; undefined when no path does. */
	exit: ReachingAssignments | undefined;
}

/** What one walk collects: each statement's reaching values, and the statements that never run. */
interface WalkOut {
	out: Map<BodyNode, ReachingAssignments>;
	dead: Set<BodyNode>;
	/** The one-line If branches a known condition decides against. */
	deadSpans: Span[];
	/** Whether Err.Raise leaves the list: not when the procedure resumes past errors. */
	raiseLeaves: boolean;
	/** The keys of the labels some GoTo, GoSub, Resume or On ... GoTo names. */
	referenced: ReadonlySet<string>;
}

/** The end of a statement list no path reaches. */
const UNREACHED: ReachingAssignments = new Map();

const WALKS = new WeakMap<readonly BodyNode[], Map<string, CachedWalk>>();

/** Each start's cache key, by identity: a kept start is asked for by several rules. */
const START_KEYS = new WeakMap<ReachingAssignments, string>();

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
	walk: WalkOut,
	caseResets = false,
): ReachingAssignments {
	let current = entry;
	for (let i = 0; i < list.length; i++) {
		const node = list[i];
		if (isInactiveNode(activity, node) || node.kind === 'VariableGroup' || node.kind === 'ConditionalDirective') {
			continue;
		}
		// A label a jump may reach starts over; one nothing names, as when
		// `On Error GoTo EH` is commented out, leaves dead code dead (issue #421).
		const label = isLeafStatement(node) ? statementLabelDeclaration(source, node.span) : undefined;
		if (label && (current !== UNREACHED || walk.referenced.has(label.key))) {
			current = NONE;
		}
		if (current === UNREACHED) {
			// After a guard that always leaves: nothing here runs (issue #273).
			markUnreachable(node, walk.dead);
			continue;
		}
		if (!isLeafStatement(node)) {
			// What holds as the block starts: a For reads its bounds here
			// (issue #200).
			record(walk.out, node, current);
			current = walkBlock(source, node, current, activity, walk);
			continue;
		}
		if (caseResets && tokenText(statementTokensAfterLeadingLabel(source, node.span)[0]) === 'case') {
			// A Select's arms are exclusive: each starts where the block did.
			current = entry;
		}
		if (node.kind === 'Statement' && node.singleLineIfBranches) {
			const group: LeafStatementNode[] = [node];
			while (i + 1 < list.length && isLeafStatement(list[i + 1]) && (list[i + 1] as LeafStatementNode).singleLineIfTail) {
				group.push(list[++i] as LeafStatementNode);
			}
			current = walkSingleLineIf(source, group, current, walk);
			continue;
		}
		record(walk.out, node, current);
		current = afterStatement(source, node.span, current);
		if (leavesTheList(source, node.span, walk.raiseLeaves)) {
			current = UNREACHED;
		}
	}
	return current;
}

/**
 * A single-line If and the statements after its colons run only with their
 * branch. With the condition known (issue #273), a false one runs nothing,
 * and a true one runs its branch in order, which may leave. Otherwise they
 * see what held before the If, less whatever the If itself changes, and so
 * does what follows.
 */
function walkSingleLineIf(
	source: string,
	group: readonly LeafStatementNode[],
	current: ReachingAssignments,
	walk: WalkOut,
): ReachingAssignments {
	const [ifStmt] = group;
	const branches = ifStmt.kind === 'Statement' ? ifStmt.singleLineIfBranches ?? [] : [];
	const condition = branches.length === 1 ? ifConditionTokens(statementTokensAfterLeadingLabel(source, ifStmt.span)) : undefined;
	const known = condition ? conditionValue(condition, factsFrom(current)) : undefined;
	if (known === false) {
		for (const stmt of group) {
			walk.dead.add(stmt);
		}
		return current;
	}
	if (known === true) {
		let state = current;
		for (const [k, stmt] of group.entries()) {
			record(walk.out, stmt, state);
			const span = k === 0 ? branches[0] : stmt.span;
			state = afterStatement(source, span, state);
			if (leavesTheList(source, span, walk.raiseLeaves)) {
				return UNREACHED;
			}
		}
		return state;
	}
	// `If x = 2 Then y = 0 Else y = Sqr(-1)` with x known: the If runs, and
	// the branch its condition decides against does not (issue #430).
	if (branches.length === 2) {
		const decided = conditionValue(ifConditionTokens(statementTokensAfterLeadingLabel(source, ifStmt.span)) ?? [], factsFrom(current));
		if (decided !== undefined) {
			const elseStart = branches[1].start;
			walk.deadSpans.push(branches[decided ? 1 : 0]);
			for (const tail of group.slice(1)) {
				if ((tail.span.start >= elseStart) === decided) {
					walk.deadSpans.push(tail.span);
				}
			}
			// One statement an arm, and no If nested in them: the arm that
			// runs is the If's whole effect, `If n > 0 Then F = 1 Else F = 0`
			// with n known leaving F known (issue #562).
			// The Then arm's span runs through the Else that ends it.
			const armToks = statementTokens(source, branches[decided ? 0 : 1]);
			const elseTok = decided ? armToks.find((tok) => tokenText(tok) === 'else') : undefined;
			const taken = elseTok ? { start: branches[0].start, end: branches[0].start + elseTok.start } : branches[decided ? 0 : 1];
			const nested = statementTokensAfterLeadingLabel(source, ifStmt.span).some((tok, k) => k > 0 && tokenText(tok) === 'if');
			if (group.length === 1 && !nested) {
				record(walk.out, ifStmt, current);
				return leavesTheList(source, taken, walk.raiseLeaves) ? UNREACHED : afterStatement(source, taken, current);
			}
		}
	}
	const touched = touchedBy(source, group);
	const after = withoutMentionedObjects(touched === 'all' ? NONE : without(current, touched), source, { start: group[0].span.start, end: group[group.length - 1].span.end });
	for (const stmt of group) {
		record(walk.out, stmt, after);
	}
	return after;
}

function walkBlock(
	source: string,
	node: BodyNode,
	entry: ReachingAssignments,
	activity: ConditionalActivityTracker | undefined,
	walk: WalkOut,
): ReachingAssignments {
	if (!('body' in node) || !Array.isArray(node.body)) {
		return entry;
	}
	// An If or a Select whose outcome is known runs one arm (issue #273).
	const chosen = knownArm(source, node, entry);
	if (chosen) {
		for (const arm of chosen.arms) {
			if (arm !== chosen.taken) {
				for (const stmt of arm) {
					markUnreachable(stmt, walk.dead);
				}
			}
		}
		return chosen.taken ? walkList(source, chosen.taken, entry, activity, walk) : entry;
	}
	// A loop known to run no pass runs none of its body (issue #406): `For i
	// = 1 To 0`, `While d <> 0` with d still 0, `For Each x In Array()`. A
	// For counter is left at its start.
	const none = loopRunsNoPass(source, node, entry);
	if (none) {
		for (const stmt of node.body as BodyNode[]) {
			markUnreachable(stmt, walk.dead);
		}
		if (none.counter === undefined) {
			return entry;
		}
		const next = new Map(entry);
		next.set(none.counter.name, rawExpressionTokens(String(none.counter.value)));
		return next;
	}
	const touched = loopTouched(source, node, entry, activity);
	const after = withoutMentionedObjects(touched === 'all' ? NONE : without(entry, touched), source, node.span);
	// An If arm, a Case or a With body runs once, from the state the block is
	// entered with; a loop's body may run again with what it changed (issue
	// #259: `If True Then x = 1 / x` reads the 0 x starts with).
	const inside = isLoopBlock(node) || touched === 'all' ? after : entry;
	if (node.kind === 'IfBlock') {
		for (const branch of (node as IfBlockNode).branches) {
			walkList(source, branch.body, inside, activity, walk);
		}
	} else {
		walkList(source, node.body as BodyNode[], inside, activity, walk, node.kind === 'SelectBlock');
	}
	const final = touched === 'all' ? undefined : forCounterFinalValue(source, node, activity) ?? doCounterFinalValue(source, node, entry, activity);
	if (final !== undefined) {
		const next = new Map(after);
		next.set(final.name, rawExpressionTokens(String(final.value)));
		return next;
	}
	return after;
}

/** The most passes the walk runs a Do loop's counter through before giving up. */
const DO_COUNTER_PASSES = 100_000;

/**
 * What a Do or While loop's counter holds after the loop ends (issue #479,
 * measured in Excel 16.0): `i = 1: Do While i <= 5: ...: i = i + 1: Loop`
 * leaves i at 6. The body steps the counter by one top-level `i = i + k` or
 * `i = i - k` and writes it nowhere else, nothing in it may leave the loop,
 * and the condition reads only the counter and names the body leaves alone.
 */
function doCounterFinalValue(
	source: string,
	node: BodyNode,
	entry: ReachingAssignments,
	activity: ConditionalActivityTracker | undefined,
): { name: string; value: number } | undefined {
	if ((node.kind !== 'DoBlock' && node.kind !== 'WhileBlock') || !('body' in node)) {
		return undefined;
	}
	const body = (node.body as BodyNode[]).filter((stmt) => !isInactiveNode(activity, stmt));
	// The condition: on the header, or on the footer of a Do.
	const header = statementTokensAfterLeadingLabel(source, blockHeaderLineSpan(source, node.span)).filter((tok) => tok.kind !== 'comment');
	const footer = node.kind === 'DoBlock' ? statementTokensAfterLeadingLabel(source, blockFooterLineSpan(source, node.span)).filter((tok) => tok.kind !== 'comment') : [];
	const headWord = tokenText(header[node.kind === 'DoBlock' ? 1 : 0]);
	const footWord = tokenText(footer[1]);
	const atHead = headWord === 'while' || headWord === 'until';
	const atFoot = footWord === 'while' || footWord === 'until';
	if (atHead === atFoot) {
		return undefined;
	}
	const condition = atHead ? header.slice(node.kind === 'DoBlock' ? 2 : 1) : footer.slice(2);
	const until = (atHead ? headWord : footWord) === 'until';
	// The one step: `i = i + k` or `i = i - k`, with k a whole number.
	let counter: string | undefined;
	let step: number | undefined;
	let stepAt: BodyNode | undefined;
	for (const stmt of body) {
		const bare = isLeafStatement(stmt) && !(stmt.kind === 'Statement' && stmt.singleLineIfBranches) ? bareAssignmentTarget(source, stmt.span) : undefined;
		const value = bare?.valueTokens.filter((tok) => tok.kind !== 'comment') ?? [];
		const lower = bare?.name.toLowerCase();
		const k = value.length === 3 && tokenName(value[0])?.toLowerCase() === lower && (value[1].rawText === '+' || value[1].rawText === '-') ? signedInteger([value[2]]) : undefined;
		if (lower && k !== undefined && literalOf(entry.get(lower)) !== undefined && conditionNames(condition).has(lower)) {
			if (counter !== undefined) {
				return undefined;
			}
			counter = lower;
			step = value[1].rawText === '+' ? k : -k;
			stepAt = stmt;
		}
	}
	const start = counter !== undefined ? literalOf(entry.get(counter)) : undefined;
	if (counter === undefined || step === undefined || step === 0 || typeof start !== 'number') {
		return undefined;
	}
	// Nothing else may write the counter or a name the condition reads, or leave.
	const rest = body.filter((stmt) => stmt !== stepAt);
	for (const lower of conditionNames(condition)) {
		if (loopBodyMayLeaveOrWrite(source, rest, lower, activity)) {
			return undefined;
		}
	}
	const facts = factsFrom(entry);
	let value = start;
	for (let pass = 0; pass <= DO_COUNTER_PASSES; pass++) {
		const known = (current: number): boolean | undefined =>
			conditionValue(condition, { value: (lower) => (lower === counter ? current : facts.value(lower)) });
		if (atHead) {
			const holds = known(value);
			if (holds === undefined) {
				return undefined;
			}
			if (holds === until) {
				return { name: counter, value };
			}
			value += step;
		} else {
			value += step;
			const holds = known(value);
			if (holds === undefined) {
				return undefined;
			}
			if (holds === until) {
				return { name: counter, value };
			}
		}
	}
	return undefined;
}

/** The names a condition reads, lowercased. */
function conditionNames(condition: readonly VbaToken[]): Set<string> {
	return mentionedNames(condition);
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

/**
 * Whether a loop runs no pass, from the state it is entered with: a For whose
 * bounds and step are known and pass each other, a `Do While` or `While`
 * whose condition is known False, a `Do Until` whose condition is known True,
 * and a For Each over `Array()`. Undefined when it may run.
 */
function loopRunsNoPass(source: string, node: BodyNode, entry: ReachingAssignments): { counter?: { name: string; value: number } } | undefined {
	if (!isLoopBlock(node)) {
		return undefined;
	}
	const toks = statementTokensAfterLeadingLabel(source, blockHeaderLineSpan(source, node.span)).filter((tok) => tok.kind !== 'comment');
	const facts = factsFrom(entry);
	if (node.kind === 'ForBlock' && node.each) {
		const inAt = toks.findIndex((tok) => tokenText(tok) === 'in');
		const group = toks.slice(inAt + 1);
		const isEmptyArray = (value: readonly VbaToken[]): boolean =>
			value.length === 3 && tokenText(value[0]) === 'array' && value[1].rawText === '(' && value[2].rawText === ')';
		// A local holding `Array()`, or a Collection nothing has added to
		// (issue #483, measured in Excel 16.0).
		const held = group.length === 1 ? entry.get(tokenName(group[0])?.toLowerCase() ?? '') : undefined;
		const empty = inAt > 0 && (isEmptyArray(group) || held === EMPTY_COLLECTION || (held !== undefined && isEmptyArray(held.filter((tok) => tok.kind !== 'comment'))));
		return empty ? {} : undefined;
	}
	if (node.kind === 'ForBlock') {
		const eq = toks.findIndex((tok) => tok.rawText === '=');
		const to = toks.findIndex((tok) => tokenText(tok) === 'to');
		const stepAt = toks.findIndex((tok) => tokenText(tok) === 'step');
		const known = (part: readonly VbaToken[]): number | undefined => {
			const literal = signedInteger(part);
			const value = literal ?? (part.length === 1 ? facts.value(tokenName(part[0])?.toLowerCase() ?? '') : undefined);
			return typeof value === 'number' ? value : undefined;
		};
		const start = eq > 0 && to > eq ? known(toks.slice(eq + 1, to)) : undefined;
		const limit = to > 0 ? known(toks.slice(to + 1, stepAt > 0 ? stepAt : toks.length)) : undefined;
		const step = stepAt > 0 ? known(toks.slice(stepAt + 1)) : 1;
		if (start === undefined || limit === undefined || step === undefined || step === 0 || !node.controlVariable) {
			return undefined;
		}
		return (step > 0 ? start > limit : start < limit) ? { counter: { name: node.controlVariable.toLowerCase(), value: start } } : undefined;
	}
	// `Do While c`, `Do Until c` and `While c`; a condition after `Loop` lets one pass run.
	const head = tokenText(toks[node.kind === 'DoBlock' ? 1 : 0]);
	if (head !== 'while' && head !== 'until') {
		return undefined;
	}
	const condition = toks.slice(node.kind === 'DoBlock' ? 2 : 1);
	const value = condition.length > 0 ? conditionValue(condition, facts) : undefined;
	return value === (head === 'until') ? {} : undefined;
}

/** The keys of every label a statement of the body jumps to or resumes at. */
function referencedLabels(source: string, body: readonly BodyNode[], activity: ConditionalActivityTracker | undefined): Set<string> {
	const out = new Set<string>();
	const visit = (list: readonly BodyNode[]): void => {
		for (const node of list) {
			if (isInactiveNode(activity, node)) {
				continue;
			}
			if (isLeafStatement(node)) {
				for (const ref of statementLabelReferences(source, node.span)) {
					out.add(ref.key);
				}
			} else if ('body' in node && Array.isArray(node.body)) {
				visit(node.body as BodyNode[]);
			}
		}
	};
	visit(body);
	return out;
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
			if ([...passedWhole(source, header, node.span.start)].includes(lower)) {
				return true;
			}
			if ('body' in node && Array.isArray(node.body) && loopBodyMayLeaveOrWrite(source, node.body as BodyNode[], lower, activity)) {
				return true;
			}
			continue;
		}
		if (jumpTargetLabelDeclaration(source, node.span)) {
			return true;
		}
		for (const span of statementAndBranchSpans(node)) {
			const toks = statementTokensAfterLeadingLabel(source, span);
			const head = tokenText(toks[0]);
			if (LEAVING_HEADS.has(head) || (head === 'end' && toks.length === 1)) {
				return true;
			}
			if (WRITING_HEADS.has(head) && writtenNames(toks).has(lower)) {
				return true;
			}
			if (bareAssignmentTarget(source, span)?.name.toLowerCase() === lower || [...passedWhole(source, toks, span.start)].includes(lower)) {
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
	// A mention of an object the walk knows may change it: `c.Add 1`.
	const known = [...mentionedNames(toks)].filter((lower) => {
		const value = before.get(lower);
		return value === OBJECT_NOTHING || value === EMPTY_COLLECTION;
	});
	before = without(before, known);
	before = without(before, [...mentionedNames(toks)].map(elementsOf));
	if (WRITING_HEADS.has(head) && !(head === 'line' && tokenText(toks[1]) !== 'input')) {
		const after = without(before, writtenNames(toks));
		const object = head === 'set' ? setObjectValue(toks) : undefined;
		if (!object) {
			return after;
		}
		const next = new Map(after);
		next.set(object.name, object.value);
		return next;
	}
	let after = without(before, passedWhole(source, toks, span.start));
	const bare = bareAssignmentTarget(source, span);
	if (bare && !identityAssignment(bare.name, bare.valueTokens)) {
		const next = new Map(after);
		const value = bare.valueTokens.filter((tok) => tok.kind !== 'comment');
		// `d = a` copies what a holds here: `a = 0: d = a` leaves d 0, and a
		// later change to a leaves d as it was (issue #346).
		const copied = value.length === 1 ? tokenName(value[0])?.toLowerCase() : undefined;
		next.set(bare.name.toLowerCase(), copied !== undefined && before.has(copied) ? before.get(copied)! : value);
		after = next;
	}
	const element = walkElements ? elementAssignment(toks, before) : undefined;
	if (element) {
		const next = new Map(after);
		next.set(element.key, element.value);
		after = next;
	}
	return after;
}

/**
 * `a(0) = Null` on an array the procedure declares, the index a whole
 * number or a local the walk knows holds one: the element's key and its
 * value (issue #332). Null is the one value a rule reads from an element.
 */
function elementAssignment(toks: readonly VbaToken[], before: ReachingAssignments): { key: string; value: readonly VbaToken[] } | undefined {
	const lower = tokenName(toks[0])?.toLowerCase();
	if (!lower || !walkArrays.has(lower) || toks[1]?.rawText !== '(') {
		return undefined;
	}
	let close = -1;
	for (let i = 2, depth = 1; i < toks.length; i++) {
		depth += toks[i].rawText === '(' ? 1 : toks[i].rawText === ')' ? -1 : 0;
		if (depth === 0) {
			close = i;
			break;
		}
	}
	if (close < 0 || toks[close + 1]?.rawText !== '=') {
		return undefined;
	}
	const index = knownIndex(toks.slice(2, close), before);
	const value = toks.slice(close + 2).filter((tok) => tok.kind !== 'comment');
	return index === undefined || value.length !== 1 || tokenText(value[0]) !== 'null' ? undefined : { key: elementKey(lower, index), value };
}

/** A subscript's value: a whole-number literal, or a local holding one here. */
export function knownIndex(subscript: readonly VbaToken[], held: ReachingAssignments): number | undefined {
	const toks = subscript.filter((tok) => tok.kind !== 'comment');
	if (toks.length !== 1) {
		return undefined;
	}
	const lower = tokenName(toks[0])?.toLowerCase();
	const value = lower !== undefined ? literalOf(held.get(lower)) : literalOf(toks);
	return typeof value === 'number' ? value : undefined;
}

/** The marker `without` reads as every element of the array: "a(". */
function elementsOf(lower: string): string {
	return `${lower}(`;
}

/** Every name a block may change, header and footer lines included, or 'all'. */
function touchedInBlock(
	source: string,
	block: BodyNode,
	activity: ConditionalActivityTracker | undefined,
	decided: (condition: readonly VbaToken[]) => boolean | undefined = () => undefined,
): Set<string> | 'all' {
	const names = new Set<string>();
	if (block.kind === 'ForBlock' && block.controlVariable) {
		names.add(block.controlVariable.toLowerCase());
	}
	for (const span of [blockHeaderLineSpan(source, block.span), blockFooterLineSpan(source, block.span)]) {
		for (const lower of passedWhole(source, statementTokensAfterLeadingLabel(source, span), span.start)) {
			names.add(lower);
		}
	}
	const visit = (list: readonly BodyNode[]): boolean => {
		for (let i = 0; i < list.length; i++) {
			const node = list[i];
			if (isInactiveNode(activity, node)) {
				continue;
			}
			if (isLeafStatement(node)) {
				// A label inside the block can be reached from anywhere, and
				// what follows the block then runs with whatever that path held.
				if (jumpTargetLabelDeclaration(source, node.span)) {
					return true;
				}
				// A one-line If with no Else that is decided False changes
				// only what its condition passes, and its tail never runs.
				if (node.kind === 'Statement' && node.singleLineIfBranches) {
					const toks = statementTokensAfterLeadingLabel(source, node.span);
					const condition = ifConditionTokens(toks);
					if (condition && !toks.some((tok) => tokenText(tok) === 'else') && decided(condition) === false) {
						for (const lower of passedWhole(source, [toks[0], ...condition], node.span.start)) {
							names.add(lower);
						}
						while (i + 1 < list.length && isLeafStatement(list[i + 1]) && (list[i + 1] as LeafStatementNode).singleLineIfTail) {
							i++;
						}
						continue;
					}
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
					// An arm decided against never runs; one decided for is the
					// last that may (issue #575).
					for (const branch of node.branches) {
						for (const lower of passedWhole(source, statementTokensAfterLeadingLabel(source, branch.headerSpan), branch.headerSpan.start)) {
							names.add(lower);
						}
						const condition = branch.branchKind === 'else' ? undefined : ifConditionTokens(statementTokensAfterLeadingLabel(source, branch.headerSpan));
						const verdict = condition ? decided(condition) : undefined;
						if (verdict === false) {
							continue;
						}
						if (visit(branch.body)) {
							return true;
						}
						if (verdict === true) {
							break;
						}
					}
					continue;
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

/**
 * The names a block may change. In a loop, a name written only in an arm that
 * what holds on every pass decides against keeps its value: with a = 0 and
 * nothing else writing a, `If a > 1 Then c.Add a` never runs, so a stays 0
 * (issue #575). The rounds start from nothing changed and add what the arms
 * still live write, until a round adds nothing: every name left out is then
 * written only where it never runs.
 */
function loopTouched(
	source: string,
	node: BodyNode,
	entry: ReachingAssignments,
	activity: ConditionalActivityTracker | undefined,
): Set<string> | 'all' {
	if (!isLoopBlock(node)) {
		return touchedInBlock(source, node, activity);
	}
	let touched = new Set<string>();
	for (;;) {
		const facts = factsFrom(without(entry, touched));
		const live = touchedInBlock(source, node, activity, (condition) => conditionValue(condition, facts));
		if (live === 'all' || [...live].every((lower) => touched.has(lower))) {
			return live === 'all' ? live : touched;
		}
		touched = new Set([...touched, ...live]);
	}
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
			const changed = WRITING_HEADS.has(head) ? writtenNames(toks) : passedWhole(source, toks, span.start);
			for (const lower of changed) {
				names.add(lower);
			}
			for (const lower of mentionedNames(toks)) {
				names.add(elementsOf(lower));
			}
			const bare = bareAssignmentTarget(source, span);
			if (bare && !identityAssignment(bare.name, bare.valueTokens)) {
				names.add(bare.name.toLowerCase());
			}
		}
	}
	return names;
}

/**
 * `d = d + 0`, `d = d * 1`, `d = 1 * d`, `d = d - 0`, `d = d / 1`: an
 * assignment that leaves a number as it was (issue #350).
 */
export function identityAssignment(name: string, valueTokens: readonly VbaToken[]): boolean {
	const value = valueTokens.filter((tok) => tok.kind !== 'comment');
	const lower = name.toLowerCase();
	const self = (tok: VbaToken | undefined): boolean => tokenName(tok)?.toLowerCase() === lower;
	const literal = (tok: VbaToken | undefined): string | undefined => (tok?.kind === 'integerLiteral' ? tok.rawText.replace(/[%&^]$/, '') : undefined);
	if (value.length !== 3) {
		return false;
	}
	const [a, op, b] = value;
	if (self(a)) {
		return ((op.rawText === '+' || op.rawText === '-') && literal(b) === '0') || ((op.rawText === '*' || op.rawText === '/') && literal(b) === '1');
	}
	return self(b) && ((op.rawText === '+' && literal(a) === '0') || (op.rawText === '*' && literal(a) === '1'));
}

/** The state less every object the walk knows that the span names: a block may have added to it or set it. */
function withoutMentionedObjects(state: ReachingAssignments, source: string, span: Span): ReachingAssignments {
	const known = [...state].filter(([, value]) => value === OBJECT_NOTHING || value === EMPTY_COLLECTION);
	if (known.length === 0) {
		return state;
	}
	const named = mentionedNames(statementTokens(source, span));
	return without(state, known.map(([lower]) => lower).filter((lower) => named.has(lower)));
}

/** `Set c = Nothing` and `Set c = New Collection`: the name and what it now holds. */
function setObjectValue(toks: readonly VbaToken[]): { name: string; value: readonly VbaToken[] } | undefined {
	const name = tokenName(toks[1])?.toLowerCase();
	if (!name || toks[2]?.rawText !== '=') {
		return undefined;
	}
	const value = toks.slice(3).filter((tok) => tok.kind !== 'comment').map((tok) => tok.rawText.toLowerCase()).join(' ');
	return value === 'nothing' ? { name, value: OBJECT_NOTHING }
		: value === 'new collection' || value === 'new vba . collection' ? { name, value: EMPTY_COLLECTION }
		: undefined;
}

/** The arrays the procedure being walked declares, by lowercased name; set while a walk runs. */
let walkArrays: ReadonlySet<string> = new Set();

/**
 * Whether the body being walked sets an element to Null, so the walk keeps
 * element keys: elsewhere `without` need not look for them, which would cost
 * a pass over every local at each statement (issue #322).
 */
let walkElements = false;

/** The names a procedure's Dim statements declare as arrays: their subscripts pass nothing. */
function localArrayNames(body: readonly BodyNode[], activity: ConditionalActivityTracker | undefined): Set<string> {
	const out = new Set<string>();
	const visit = (list: readonly BodyNode[]): void => {
		for (const node of list) {
			if (isInactiveNode(activity, node)) {
				continue;
			}
			if (node.kind === 'VariableGroup' && !node.isConst) {
				for (const decl of node.declarations) {
					if (decl.isArray) {
						out.add(decl.name.toLowerCase());
					}
				}
			} else if (node.kind === 'IfBlock') {
				for (const branch of (node as IfBlockNode).branches) {
					visit(branch.body);
				}
			} else if ('body' in node && Array.isArray(node.body)) {
				visit(node.body as BodyNode[]);
			}
		}
	};
	visit(body);
	return out;
}

function passedWhole(source: string, toks: readonly VbaToken[], spanStart: number): Iterable<string> {
	const hits = trackedLocalsNamedWhole(toks, spanStart, () => true, READ_ONLY_INTRINSICS, walkArrays, calleeKeepsArgument(source));
	// A VBA library function assigns none of its arguments: `Left$("abc", n)`
	// leaves n as it was (issue #565).
	for (const [lower, at] of hits) {
		const index = toks.findIndex((tok) => spanStart + tok.start === at);
		if (index >= 0 && libraryFunctionArgument(toks, index)) {
			hits.delete(lower);
		}
	}
	return hits.keys();
}

/**
 * The names a writing statement may change: a ReDim's arrays, not the
 * names its bounds read, `ReDim a(n)` leaving n as it was (issue #350);
 * every name of any other.
 */
function writtenNames(toks: readonly VbaToken[]): Set<string> {
	if (tokenText(toks[0]) !== 'redim') {
		return mentionedNames(toks);
	}
	const start = tokenText(toks[1]) === 'preserve' ? 2 : 1;
	const names = new Set<string>();
	for (const group of splitTopLevelTokenGroups(toks, start, ',', toks.length)) {
		const lower = tokenName(group.find((tok) => tok.kind !== 'comment'))?.toLowerCase();
		if (lower) {
			names.add(lower);
		}
	}
	return names;
}

/** The procedures the module being walked declares, by lowercased name; set while a walk runs. */
let walkProcedures: ReadonlySet<string> = new Set();

/** The Subs, Functions, Properties and Declares a module's source declares: one of them hides a VBA function. */
function moduleProcedureNames(source: string): ReadonlySet<string> {
	// One module is walked many times in a row; keeping only the last spares
	// a cache that grows with every edit.
	if (lastProcedures?.source !== source) {
		lastProcedures = {
			source,
			names: new Set([...source.matchAll(/^[ \t]*(?:(?:Public|Private|Friend|Static|Global)[ \t]+)*(?:Sub|Function|Property[ \t]+(?:Get|Let|Set)|Declare(?:[ \t]+PtrSafe)?[ \t]+(?:Sub|Function))[ \t]+([A-Za-z_]\w*)/gim)].map((match) => match[1].toLowerCase())),
		};
	}
	return lastProcedures.names;
}

let lastProcedures: { source: string; names: ReadonlySet<string> } | undefined;

/** Whether the name at `at` stands in the parentheses of a VBA library function's call. */
function libraryFunctionArgument(toks: readonly VbaToken[], at: number): boolean {
	let depth = 0;
	for (let j = at - 1; j > 0; j--) {
		if (toks[j].rawText === ')') {
			depth++;
		} else if (toks[j].rawText === '(' && depth-- === 0) {
			// `Left$(` lexes as Left and a `$`.
			const callAt = toks[j - 1]?.rawText === '$' ? j - 2 : j - 1;
			const callee = tokenName(toks[callAt])?.toLowerCase();
			const qualified = toks[callAt - 1]?.rawText === '.';
			if (!callee || (qualified && tokenText(toks[callAt - 2]) !== 'vba') || (!qualified && walkProcedures.has(callee))) {
				return false;
			}
			return (resolveRuntimeFunction(callee) ?? resolveRuntimeFunction(`${callee}$`))?.kind === 'function';
		}
	}
	return false;
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

/**
 * The map less each name. A name drops its array's elements too, and a
 * name ending "(" drops only the elements: "a(" drops "a(0)".
 */
function without(map: ReachingAssignments, names: Iterable<string>): ReachingAssignments {
	let next: Map<string, readonly VbaToken[]> | undefined;
	const elements = walkElements ? [...map.keys()].filter((key) => key.endsWith(')')) : [];
	for (const lower of names) {
		if ((next ?? map).has(lower)) {
			next ??= new Map(map);
			next.delete(lower);
		}
		if (elements.length > 0) {
			const prefix = lower.endsWith('(') ? lower : elementsOf(lower);
			for (const key of elements) {
				if (key.startsWith(prefix) && (next ?? map).has(key)) {
					next ??= new Map(map);
					next.delete(key);
				}
			}
		}
	}
	return next ?? map;
}

/** The numbers and strings the reaching assignments give their names, for a condition. */
function factsFrom(current: ReachingAssignments): ConditionFacts {
	return {
		value: (lower) => literalOf(current.get(lower)),
		isNothing: (lower) => (current.get(lower) === OBJECT_NOTHING ? true : current.get(lower) === EMPTY_COLLECTION ? false : undefined),
		range: (lower) => datePartRange(current.get(lower)),
	};
}

/** What VBA's date-part functions return: Second(Now) is 0 to 59. */
const DATE_PART_RANGES: Readonly<Record<string, readonly [number, number]>> = {
	second: [0, 59], minute: [0, 59], hour: [0, 23], day: [1, 31], month: [1, 12], weekday: [1, 7],
};

/**
 * The range a value lies in where it is a date part plus or minus whole
 * numbers: `Second(Now) + 1000` is 1000 to 1059 (issue #565, measured in
 * Excel 16.0: `If b > 5000` is then False on every run).
 */
function datePartRange(value: readonly VbaToken[] | undefined): readonly [number, number] | undefined {
	const toks = (value ?? []).filter((tok) => tok.kind !== 'comment');
	let range: [number, number] | undefined;
	let sign = 1;
	for (let i = 0; i < toks.length;) {
		const word = tokenText(toks[i]);
		let part: readonly [number, number] | undefined;
		if (toks[i].kind === 'integerLiteral') {
			const n = parseVbaIntegerLiteral(toks[i].rawText);
			if (n === undefined) {
				return undefined;
			}
			part = [n, n];
			i++;
		} else if (DATE_PART_RANGES[word] && toks[i + 1]?.rawText === '(' && toks[i - 1]?.rawText !== '.') {
			const close = matchParenFrom([...toks], i + 1);
			if (close < 0) {
				return undefined;
			}
			part = DATE_PART_RANGES[word];
			i = close + 1;
		} else {
			return undefined;
		}
		const next: [number, number] = sign > 0 ? [part[0], part[1]] : [-part[1], -part[0]];
		range = range ? [range[0] + next[0], range[1] + next[1]] : next;
		if (i >= toks.length) {
			break;
		}
		if (toks[i].rawText !== '+' && toks[i].rawText !== '-') {
			return undefined;
		}
		sign = toks[i].rawText === '+' ? 1 : -1;
		i++;
	}
	// A lone number is a literal, which the value holds exactly.
	return range && toks.some((tok) => DATE_PART_RANGES[tokenText(tok)]) ? range : undefined;
}

/** A value's tokens as one number or string literal, a sign allowed. */
function literalOf(value: readonly VbaToken[] | undefined): number | string | undefined {
	const toks = (value ?? []).filter((tok) => tok.kind !== 'comment');
	const negative = toks.length === 2 && toks[0].rawText === '-';
	const tok = toks[negative ? 1 : 0];
	if (!tok || toks.length !== (negative ? 2 : 1)) {
		return undefined;
	}
	if (tok.kind === 'integerLiteral') {
		const number = parseVbaIntegerLiteral(tok.rawText);
		return number === undefined ? undefined : negative ? -number : number;
	}
	return tok.kind === 'stringLiteral' && !negative ? tok.rawText.slice(1, -1).replace(/""/g, '"') : undefined;
}

/**
 * The arm of an If or a Select that runs when the walk knows its outcome:
 * every arm, and the one taken (undefined when none is). Undefined when
 * the outcome is not known.
 */
function knownArm(source: string, node: BodyNode, entry: ReachingAssignments): { arms: readonly (readonly BodyNode[])[]; taken: readonly BodyNode[] | undefined } | undefined {
	const facts = factsFrom(entry);
	if (node.kind === 'IfBlock') {
		const arms = node.branches.map((branch) => branch.body);
		for (const branch of node.branches) {
			if (branch.branchKind === 'else') {
				return { arms, taken: branch.body };
			}
			const condition = ifConditionTokens(statementTokensAfterLeadingLabel(source, branch.headerSpan));
			const known = condition ? conditionValue(condition, facts) : undefined;
			if (known === undefined) {
				return undefined;
			}
			if (known) {
				return { arms, taken: branch.body };
			}
		}
		return { arms, taken: undefined };
	}
	if (node.kind !== 'SelectBlock') {
		return undefined;
	}
	// `Select Case d` with d known: the first Case whose values match.
	const header = statementTokensAfterLeadingLabel(source, blockHeaderLineSpan(source, node.span)).filter((tok) => tok.kind !== 'comment');
	const selector = header.length === 3 && tokenText(header[1]) === 'case' ? literalOf(header.slice(2)) ?? factsFrom(entry).value(tokenName(header[2])?.toLowerCase() ?? '') : undefined;
	if (typeof selector !== 'number') {
		return undefined;
	}
	const arms = selectArms(source, node.body as BodyNode[]);
	for (const arm of arms) {
		const caseLine = arm.find((stmt) => isLeafStatement(stmt) && tokenText(statementTokensAfterLeadingLabel(source, stmt.span)[0]) === 'case');
		if (!caseLine) {
			continue;
		}
		const matched = caseMatches(statementTokensAfterLeadingLabel(source, caseLine.span).filter((tok) => tok.kind !== 'comment'), selector);
		if (matched === undefined) {
			return undefined;
		}
		if (matched) {
			return { arms, taken: arm };
		}
	}
	return { arms, taken: undefined };
}

/** Whether a `Case` line's values take a number: literals, `Is op n`, `a To b`, Else. */
function caseMatches(toks: readonly VbaToken[], selector: number): boolean | undefined {
	if (tokenText(toks[1]) === 'else') {
		return true;
	}
	let anyUnknown = false;
	for (const item of splitTopLevelTokenGroups(toks.slice(1), 0, ',')) {
		let matched: boolean | undefined;
		const to = item.findIndex((tok) => tokenText(tok) === 'to');
		if (tokenText(item[0]) === 'is' && item[1]?.kind === 'operator') {
			const value = literalOf(item.slice(2));
			matched = typeof value === 'number' ? conditionValue(rawExpressionTokens(`${selector} ${item[1].rawText} ${value}`), { value: () => undefined }) : undefined;
		} else if (to > 0) {
			const low = literalOf(item.slice(0, to));
			const high = literalOf(item.slice(to + 1));
			matched = typeof low === 'number' && typeof high === 'number' ? selector >= low && selector <= high : undefined;
		} else {
			const value = literalOf(item);
			matched = typeof value === 'number' ? value === selector : undefined;
		}
		if (matched === true) {
			return true;
		}
		if (matched === undefined) {
			anyUnknown = true;
		}
	}
	return anyUnknown ? undefined : false;
}

/** Marks a statement, and every statement inside it, as never running. */
function markUnreachable(node: BodyNode, dead: Set<BodyNode>): void {
	dead.add(node);
	if (node.kind === 'IfBlock') {
		for (const branch of node.branches) {
			for (const stmt of branch.body) {
				markUnreachable(stmt, dead);
			}
		}
		return;
	}
	if ('body' in node && Array.isArray(node.body)) {
		for (const stmt of node.body as BodyNode[]) {
			markUnreachable(stmt, dead);
		}
	}
}

function record(out: Map<BodyNode, ReachingAssignments>, stmt: BodyNode, current: ReachingAssignments): void {
	if (current.size > 0) {
		out.set(stmt, current);
	}
}
