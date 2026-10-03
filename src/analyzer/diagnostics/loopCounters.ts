// Loop counters and the values their passes take (issue #200).
//
// The usual off-by-one bug is a counter one step past what it indexes:
// `For i = 0 To Len(s) - 1` into Mid$, which starts at 1, or
// `For i = LBound(a) To UBound(a) + 1` into a. The value rules check a
// statement once as written; here each statement of a loop that runs on
// every pass learns the counter's first and last values, so a rule can check
// the statement again with the counter bound to each.
//
// A counter is a For counter, or a Do/While counter the code steps itself:
// `i = 1: Do While i <= 4 ... i = i + 1: Loop`. Its bounds are a whole number,
// or Len, UBound, LBound or .Count of a name, plus or minus a whole number,
// so `UBound(a) + 1` is known to pass a's last element whatever a holds.
//
// A statement runs on every pass when it sits in the loop's own body, or in
// a With there, before anything that may leave the pass: Exit, GoTo, Resume,
// Return, End, Stop, a label (which a GoTo may reach), or a block holding
// any of those. A statement inside an If or an inner loop may not run on the
// first or the last pass, and is left alone. A loop whose body writes the
// counter has no counter here, and a bound whose name the body writes (an
// assignment, ReDim, a ByRef argument, Add or Remove) is not read: For reads
// its bounds once, before the first pass.

import type { ConditionalActivityTracker } from '../conditional/conditionalCompilation';
import { parseVbaIntegerLiteral } from '../constants/integerConstantExpression';
import { jumpTargetLabelDeclaration } from '../flow/procedureLabels';
import { splitTopLevelTokenGroups } from '../lexer/tokenHelpers';
import type { VbaToken } from '../lexer/tokenKinds';
import type { BodyNode, LeafStatementNode, Span } from '../parser/nodes';
import { isLeafStatement } from '../parser/nodes';
import type { PushFn } from './analysisContext';
import { stringLiteralValue } from './typeInference';
import {
	bareAssignmentTarget,
	blockHeaderLineSpan,
	isInactiveNode,
	matchParenFrom,
	setAssignmentTarget,
	statementAndBranchSpans,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
} from './walker';

/** A bound's named part: `Len(s)`, `UBound(a)`, `LBound(a, 2)`, `c.Count`. */
export interface CounterAtom {
	kind: 'len' | 'ubound' | 'lbound' | 'count';
	/** The name it reads, lower-cased. */
	name: string;
	/** As written, for messages. */
	text: string;
	/** The dimension UBound or LBound asks for; 1 otherwise. */
	dimension: number;
}

/** A counter value: the atom plus the offset, or the offset alone. */
export interface CounterValue {
	atom?: CounterAtom;
	offset: number;
}

export interface LoopCounter {
	/** As the header writes it. */
	name: string;
	first: CounterValue;
	/** Undefined when the last pass cannot be told: a symbolic bound with a Step other than 1 or -1. */
	last: CounterValue | undefined;
	step: number;
	/** 'For', or 'Do' for a counter the loop steps itself. */
	loop: 'For' | 'Do';
	/** The loop, whose bounds read the values that hold as it starts. */
	loopNode: BodyNode;
}

/** The counter in force at a statement that runs on every pass, by lower-cased name. */
export type CountersAt = ReadonlyMap<string, LoopCounter>;

/** Statement heads that write every name they mention. */
const WRITING_HEADS: ReadonlySet<string> = new Set(['set', 'redim', 'erase', 'input', 'get', 'line', 'lset', 'rset', 'mid', 'mid$']);

/** Statement heads after which the rest of the pass may not run. */
const LEAVING_HEADS: ReadonlySet<string> = new Set(['exit', 'goto', 'gosub', 'resume', 'return', 'end', 'stop']);

const WALKS = new WeakMap<readonly BodyNode[], {
	source: string;
	activity: ConditionalActivityTracker | undefined;
	result: ReadonlyMap<LeafStatementNode, CountersAt>;
}>();

/** For each statement that runs on every pass of a loop with a counter, that counter. */
export function loopCountersAt(
	source: string,
	body: readonly BodyNode[],
	activity: ConditionalActivityTracker | undefined,
): ReadonlyMap<LeafStatementNode, CountersAt> {
	const cached = WALKS.get(body);
	if (cached && cached.source === source && cached.activity === activity) {
		return cached.result;
	}
	const out = new Map<LeafStatementNode, CountersAt>();
	const visit = (nodes: readonly BodyNode[]): void => {
		for (let k = 0; k < nodes.length; k++) {
			const node = nodes[k];
			if (isInactiveNode(activity, node) || !('body' in node) || !Array.isArray(node.body)) {
				continue;
			}
			const loopBody = node.body as BodyNode[];
			const found = node.kind === 'ForBlock'
				? forCounter(source, node.span, node.each, node.controlVariable, loopBody, activity)
				: node.kind === 'DoBlock' || node.kind === 'WhileBlock'
					? steppedCounter(source, node.span, nodes[k - 1], loopBody, activity)
					: undefined;
			if (found) {
				const counter: LoopCounter = { ...found.counter, loopNode: node };
				const counters = new Map([[counter.name.toLowerCase(), counter]]);
				const leaves: LeafStatementNode[] = [];
				everyPassLeaves(source, found.body, activity, leaves);
				for (const leaf of leaves) {
					out.set(leaf, counters);
				}
			}
			visit(loopBody);
		}
	};
	visit(body);
	if (out.size === 0) {
		// Most procedures have no counter, and the walk that says so costs
		// less than remembering it: a WeakMap entry per procedure body cost
		// six times the walk over the differential corpus (issue #200).
		return NO_COUNTERS;
	}
	WALKS.set(body, { source, activity, result: out });
	return out;
}

const NO_COUNTERS: ReadonlyMap<LeafStatementNode, CountersAt> = new Map();

/** `For i = <bound> To <bound> [Step <whole number>]`. */
function forCounter(
	source: string,
	span: Span,
	each: boolean,
	controlVariable: string | undefined,
	body: readonly BodyNode[],
	activity: ConditionalActivityTracker | undefined,
): { counter: Omit<LoopCounter, 'loopNode'>; body: readonly BodyNode[] } | undefined {
	if (each || !controlVariable) {
		return undefined;
	}
	const toks = statementTokensAfterLeadingLabel(source, blockHeaderLineSpan(source, span)).filter((tok) => tok.kind !== 'comment');
	const eq = toks.findIndex((tok) => tok.rawText === '=');
	const to = topLevelWordIndex(toks, 'to', eq + 1);
	if (eq < 0 || to < 0) {
		return undefined;
	}
	const stepAt = topLevelWordIndex(toks, 'step', to + 1);
	const step = stepAt < 0 ? 1 : wholeNumber(toks.slice(stepAt + 1));
	if (step === undefined || step === 0) {
		return undefined;
	}
	// Most bounds are no whole number and no atom, `.End(xlUp).Row`: read them
	// before walking the body.
	const from = counterValue(toks.slice(eq + 1, to));
	const bound = from && counterValue(toks.slice(to + 1, stepAt < 0 ? toks.length : stepAt));
	if (!from || !bound) {
		return undefined;
	}
	const written = namesWrittenIn(source, body, activity);
	const first = readable(from, written);
	const upTo = readable(bound, written);
	if (written.has(controlVariable.toLowerCase()) || !first || !upTo) {
		return undefined;
	}
	let last: CounterValue | undefined = upTo;
	if (Math.abs(step) !== 1) {
		// The last pass is the last first + k*step not past the bound.
		last = first.atom || upTo.atom
			? undefined
			: { offset: first.offset + Math.floor((upTo.offset - first.offset) / step) * step };
	}
	return { counter: { name: controlVariable, first, last, step, loop: 'For' }, body };
}

/**
 * `i = <bound>` just before `Do While i <= <bound>` (or `<`, or `Do Until i >`
 * or `>=`, or `While`), with `i = i + 1` the loop's last statement and no
 * other write to i in it.
 */
function steppedCounter(
	source: string,
	span: Span,
	before: BodyNode | undefined,
	body: readonly BodyNode[],
	activity: ConditionalActivityTracker | undefined,
): { counter: Omit<LoopCounter, 'loopNode'>; body: readonly BodyNode[] } | undefined {
	if (!before || !isLeafStatement(before)) {
		return undefined;
	}
	const init = bareAssignmentTarget(source, before.span);
	if (!init) {
		return undefined;
	}
	const lower = init.name.toLowerCase();
	const header = statementTokensAfterLeadingLabel(source, blockHeaderLineSpan(source, span)).filter((tok) => tok.kind !== 'comment');
	let i = tokenText(header[0]) === 'do' ? 1 : tokenText(header[0]) === 'while' ? 0 : -1;
	const test = tokenText(header[i]);
	if (i < 0 || (test !== 'while' && test !== 'until') || tokenName(header[i + 1])?.toLowerCase() !== lower) {
		return undefined;
	}
	i += 2;
	let operator = header[i]?.rawText ?? '';
	let next = i + 1;
	if ((operator === '<' || operator === '>') && header[i + 1]?.rawText === '=') {
		operator += '=';
		next++;
	}
	// The last value the test lets in.
	const shift = test === 'while'
		? operator === '<=' ? 0 : operator === '<' ? -1 : undefined
		: operator === '>' ? 0 : operator === '>=' ? -1 : undefined;
	if (shift === undefined) {
		return undefined;
	}
	const passes = body.filter((node) => !isInactiveNode(activity, node) && node.kind !== 'VariableGroup');
	const increment = passes[passes.length - 1];
	if (!increment || !isLeafStatement(increment) || !isIncrementOf(source, increment.span, lower)) {
		return undefined;
	}
	const rest = passes.slice(0, -1);
	const written = namesWrittenIn(source, rest, activity);
	const first = counterValue(init.valueTokens);
	const limit = readable(counterValue(header.slice(next)), written);
	if (written.has(lower) || !first || first.atom || !limit) {
		return undefined;
	}
	return {
		counter: { name: init.name, first, last: { atom: limit.atom, offset: limit.offset + shift }, step: 1, loop: 'Do' },
		body: rest,
	};
}

/** `i = i + 1`. */
function isIncrementOf(source: string, span: Span, lower: string): boolean {
	const target = bareAssignmentTarget(source, span);
	const value = target?.valueTokens.filter((tok) => tok.kind !== 'comment') ?? [];
	return target?.name.toLowerCase() === lower && value.length === 3
		&& tokenName(value[0])?.toLowerCase() === lower && value[1].rawText === '+'
		&& value[2].kind === 'integerLiteral' && parseVbaIntegerLiteral(value[2].rawText) === 1;
}

function topLevelWordIndex(toks: readonly VbaToken[], word: string, from: number): number {
	let depth = 0;
	for (let i = Math.max(0, from); i < toks.length; i++) {
		const raw = toks[i].rawText;
		if (raw === '(') {
			depth++;
		} else if (raw === ')') {
			depth--;
		} else if (depth === 0 && tokenText(toks[i]) === word) {
			return i;
		}
	}
	return -1;
}

function wholeNumber(toks: readonly VbaToken[]): number | undefined {
	if (toks.length === 1 && toks[0].kind === 'integerLiteral') {
		return parseVbaIntegerLiteral(toks[0].rawText);
	}
	if (toks.length === 2 && (toks[0].rawText === '-' || toks[0].rawText === '+') && toks[1].kind === 'integerLiteral') {
		const value = parseVbaIntegerLiteral(toks[1].rawText);
		return value === undefined ? undefined : toks[0].rawText === '-' ? -value : value;
	}
	return undefined;
}

/** A bound: a whole number, or an atom plus or minus one. */
export function counterValue(tokens: readonly VbaToken[]): CounterValue | undefined {
	const toks = tokens.filter((tok) => tok.kind !== 'comment');
	const whole = wholeNumber(toks);
	if (whole !== undefined) {
		return { offset: whole };
	}
	let core = toks;
	let offset = 0;
	const sign = toks[toks.length - 2]?.rawText;
	if (toks.length >= 3 && (sign === '+' || sign === '-') && toks[toks.length - 1].kind === 'integerLiteral') {
		const value = parseVbaIntegerLiteral(toks[toks.length - 1].rawText);
		if (value === undefined) {
			return undefined;
		}
		offset = sign === '-' ? -value : value;
		core = toks.slice(0, -2);
	}
	// `c.Count`
	if (core.length === 3 && tokenName(core[0]) && core[1].rawText === '.' && tokenText(core[2]) === 'count') {
		return { atom: { kind: 'count', name: tokenName(core[0])!.toLowerCase(), text: `${core[0].rawText}.Count`, dimension: 1 }, offset };
	}
	const callee = tokenText(core[0]);
	if ((callee !== 'len' && callee !== 'ubound' && callee !== 'lbound') || core[1]?.rawText !== '(' || matchParenFrom(core, 1) !== core.length - 1) {
		return undefined;
	}
	const args = splitTopLevelTokenGroups(core, 2, ',', core.length - 1);
	if (callee === 'len' && args.length === 1 && args[0].length === 1 && args[0][0].kind === 'stringLiteral') {
		return { offset: stringLiteralValue(args[0][0].rawText).length + offset };
	}
	const name = args[0]?.length === 1 ? tokenName(args[0][0]) : undefined;
	const dimension = args.length === 2 ? wholeNumber(args[1]) : 1;
	if (!name || dimension === undefined || args.length > (callee === 'len' ? 1 : 2)) {
		return undefined;
	}
	const text = `${core[0].rawText}(${name}${args.length === 2 ? `, ${dimension}` : ''})`;
	return { atom: { kind: callee, name: name.toLowerCase(), text, dimension }, offset };
}

/** The value, unless the loop body writes the name its atom reads. */
function readable(value: CounterValue | undefined, written: ReadonlySet<string>): CounterValue | undefined {
	return value?.atom && written.has(value.atom.name) ? undefined : value;
}

/**
 * The names a body may write: assignment and Set targets, every name a
 * writing statement mentions, bare arguments of a call (ByRef), receivers of
 * Add, Remove and Clear, and the counters of loops inside it.
 */
function namesWrittenIn(source: string, body: readonly BodyNode[], activity: ConditionalActivityTracker | undefined): Set<string> {
	const out = new Set<string>();
	const visit = (nodes: readonly BodyNode[]): void => {
		for (const node of nodes) {
			if (isInactiveNode(activity, node)) {
				continue;
			}
			if (node.kind === 'ForBlock' && node.controlVariable) {
				out.add(node.controlVariable.toLowerCase());
			}
			if ('body' in node && Array.isArray(node.body)) {
				visit(node.body as BodyNode[]);
				continue;
			}
			if (!isLeafStatement(node)) {
				continue;
			}
			for (const span of statementAndBranchSpans(node)) {
				const toks = statementTokensAfterLeadingLabel(source, span);
				const target = bareAssignmentTarget(source, span) ?? setAssignmentTarget(source, span);
				if (target) {
					out.add(target.name.toLowerCase());
				}
				// `Mid(s, i, 1) = "x"` writes s alone: its start and length are
				// read (issue #327).
				const head = tokenText(toks[0]) ?? '';
				if (['mid', 'mid$', 'midb', 'midb$'].includes(head) && (toks[1]?.rawText === '(' || (toks[1]?.rawText === '$' && toks[2]?.rawText === '('))) {
					const lower = tokenName(toks[toks[1].rawText === '(' ? 2 : 3])?.toLowerCase();
					if (lower) {
						out.add(lower);
					}
				} else if (WRITING_HEADS.has(head)) {
					for (const tok of toks) {
						const lower = tokenName(tok)?.toLowerCase();
						if (lower) {
							out.add(lower);
						}
					}
				}
				for (const lower of callStatementArguments(toks)) {
					out.add(lower);
				}
				for (let i = 0; i + 2 < toks.length; i++) {
					const member = tokenText(toks[i + 2]);
					if (toks[i + 1].rawText === '.' && (member === 'add' || member === 'remove' || member === 'clear')) {
						const lower = tokenName(toks[i])?.toLowerCase();
						if (lower) {
							out.add(lower);
						}
					}
				}
			}
		}
	};
	visit(body);
	return out;
}

/**
 * The names a call statement passes whole, which it may change ByRef:
 * `Bump i`, `Call Bump(i)`, `obj.Move i`. A name inside an expression's
 * parentheses, `a(i)` or `Mid$(s, i, 1)`, is read: a loop counter changed
 * through a function's ByRef argument is rare enough to leave.
 */
function callStatementArguments(toks: readonly VbaToken[]): string[] {
	const call = tokenText(toks[0]) === 'call';
	const head = call ? 1 : 0;
	if (!tokenName(toks[head]) && toks[head]?.rawText !== '.') {
		return [];
	}
	const out: string[] = [];
	let depth = 0;
	for (let i = head + 1; i < toks.length; i++) {
		const raw = toks[i].rawText;
		if (raw === '(') {
			depth++;
			continue;
		}
		if (raw === ')') {
			depth--;
			continue;
		}
		if (depth === 0 && toks[i].kind === 'operator' && raw === '=') {
			return []; // an assignment, not a call
		}
		const lower = tokenName(toks[i])?.toLowerCase();
		const whole = toks[i - 1]?.rawText !== '.' && toks[i + 1]?.rawText !== '(' && toks[i + 1]?.rawText !== '.';
		if (lower && whole && depth === (call ? 1 : 0)) {
			out.push(lower);
		}
	}
	return out;
}

/**
 * Collects the leaves of a loop body that run on every pass; false once
 * something may end the pass early, so nothing after it does.
 */
function everyPassLeaves(
	source: string,
	body: readonly BodyNode[],
	activity: ConditionalActivityTracker | undefined,
	out: LeafStatementNode[],
): boolean {
	for (const node of body) {
		if (isInactiveNode(activity, node) || node.kind === 'VariableGroup' || node.kind === 'ConditionalDirective') {
			continue;
		}
		if (isLeafStatement(node)) {
			if (jumpTargetLabelDeclaration(source, node.span) || mayLeave(source, node)) {
				return false;
			}
			if (node.kind !== 'Statement' || !node.singleLineIfBranches) {
				out.push(node);
			}
			continue;
		}
		if (node.kind === 'WithBlock') {
			if (!everyPassLeaves(source, node.body, activity, out)) {
				return false;
			}
			continue;
		}
		if ('body' in node && Array.isArray(node.body) && blockMayLeave(source, node.body as BodyNode[])) {
			return false;
		}
	}
	return true;
}

function mayLeave(source: string, node: LeafStatementNode): boolean {
	return statementAndBranchSpans(node).some((span) => {
		const toks = statementTokensAfterLeadingLabel(source, span);
		const head = tokenText(toks[0]) ?? '';
		return LEAVING_HEADS.has(head) || (head === 'err' && tokenText(toks[2]) === 'raise');
	});
}

function blockMayLeave(source: string, body: readonly BodyNode[]): boolean {
	return body.some((node) => isLeafStatement(node)
		? jumpTargetLabelDeclaration(source, node.span) !== undefined || mayLeave(source, node)
		: 'body' in node && Array.isArray(node.body) && blockMayLeave(source, node.body as BodyNode[]));
}

/** The number a counter value is, given what each atom is; undefined when an atom is unknown. */
export function counterNumber(
	value: CounterValue | undefined,
	counter: LoopCounter,
	atomValue: (atom: CounterAtom, counter: LoopCounter) => number | undefined,
): number | undefined {
	if (!value) {
		return undefined;
	}
	if (!value.atom) {
		return value.offset;
	}
	const base = atomValue(value.atom, counter);
	return base === undefined ? undefined : base + value.offset;
}

/** `UBound(a) + 1`, `0`, `c.Count - 1`: a counter value as the code would write it. */
export function counterText(value: CounterValue): string {
	if (!value.atom) {
		return String(value.offset);
	}
	if (value.offset === 0) {
		return value.atom.text;
	}
	return `${value.atom.text} ${value.offset > 0 ? '+' : '-'} ${Math.abs(value.offset)}`;
}

export interface CounterPass {
	pass: 'first' | 'last';
	counter: LoopCounter;
	value: number;
}

/**
 * The passes whose value is a number, given the atoms a rule knows. None
 * when the numbers show the loop never runs: `For i = 0 To Len("") - 1`.
 */
export function numericCounterPasses(
	counter: LoopCounter,
	atomValue: (atom: CounterAtom, counter: LoopCounter) => number | undefined,
): CounterPass[] {
	const first = counterNumber(counter.first, counter, atomValue);
	const last = counterNumber(counter.last, counter, atomValue);
	if (first !== undefined && last !== undefined && (counter.step > 0 ? first > last : first < last)) {
		return [];
	}
	const out: CounterPass[] = [];
	if (first !== undefined) {
		out.push({ pass: 'first', counter, value: first });
	}
	if (last !== undefined && last !== first) {
		out.push({ pass: 'last', counter, value: last });
	}
	return out;
}

/**
 * Runs a statement check as written, then once with the counter bound to
 * each pass's value, and reports what only a pass makes fail, with the pass
 * named. `check` reads the bound value for the counter's name in `values`.
 */
export function checkEachCounterPass(
	source: string,
	span: Span,
	counters: CountersAt | undefined,
	atomValue: (atom: CounterAtom, counter: LoopCounter) => number | undefined,
	check: (values: ReadonlyMap<string, number>, push: PushFn) => void,
	push: PushFn,
): void {
	// Most statements sit in no loop, or never read its counter: check them
	// once, as written, with nothing allocated.
	const passes = counters ? namedCounterPasses(source, span, counters, atomValue) : NO_PASSES;
	if (passes.length === 0) {
		check(NO_VALUES, push);
		return;
	}
	const seen = new Set<string>();
	const keyOf = (at: Span): string => `${at.start}:${at.end}`;
	check(NO_VALUES, (rule, message, at, data) => {
		seen.add(keyOf(at));
		push(rule, message, at, data);
	});
	for (const pass of passes) {
		const { counter } = pass;
		check(new Map([[counter.name.toLowerCase(), pass.value]]), (rule, message, at, data) => {
			const key = keyOf(at);
			if (seen.has(key)) {
				return;
			}
			seen.add(key);
			push(rule, `On the ${pass.pass} pass of the ${counter.loop} loop, where '${counter.name}' is ${pass.value}: ${message}`, at, data);
		});
	}
}

const NO_VALUES: ReadonlyMap<string, number> = new Map();
const NO_PASSES: readonly CounterPass[] = [];

/** The numeric passes of the counters a statement names. */
function namedCounterPasses(
	source: string,
	span: Span,
	counters: CountersAt,
	atomValue: (atom: CounterAtom, counter: LoopCounter) => number | undefined,
): CounterPass[] {
	const text = source.slice(span.start, span.end).toLowerCase();
	const out: CounterPass[] = [];
	for (const counter of counters.values()) {
		if (mentions(text, counter.name.toLowerCase())) {
			out.push(...numericCounterPasses(counter, atomValue));
		}
	}
	return out;
}

const IDENTIFIER_CHARACTER = /[A-Za-z0-9_$\u00C0-\uFFFF]/;

/** Whether lower-cased text names a lower-cased identifier as a whole word. */
function mentions(text: string, name: string): boolean {
	for (let at = text.indexOf(name); at >= 0; at = text.indexOf(name, at + 1)) {
		const before = text[at - 1] ?? ' ';
		const after = text[at + name.length] ?? ' ';
		if (!IDENTIFIER_CHARACTER.test(before) && !IDENTIFIER_CHARACTER.test(after)) {
			return true;
		}
	}
	return false;
}
