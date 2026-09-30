// Rule family: a VBA.Collection whose contents the code makes plain (issue #121).
//
// Every case was measured in Excel 16.0 (build 20326, 2026-09-26) on a local
// `New Collection` with only the Adds shown; each compiles and raises every
// time it runs.
//
//  - collection-index-out-of-range: `c(1)`, `c(0)`, `c(-1)` or `c.Remove 1`
//    with nothing added -> 5 (Invalid procedure call or argument); `c(0)`,
//    `c(-1)`, `c.Item(2)`, `c(3)` after two Adds, `c.Remove 0`, `c.Remove 2`
//    after one Add -> 9 (Subscript out of range). Collections are 1-based.
//  - collection-key-not-found: `c("nokey")`, `c.Item("nokey")`, `c.Remove "x"`
//    when the key was never added, or was removed -> 5. Keys compare without
//    case: `c.Add 1, "k"` then `c("K")` runs.
//  - collection-key-in-use: `c.Add 1, "k"` then `c.Add 2, "K"` -> 457.
//
// The rule follows a procedure's top-level statements in order, as the file
// rule does: a block ends what is known, and any use of the variable other
// than Add, Remove, Item, Count and indexing ends it too.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import { parseVbaIntegerLiteral } from '../../constants/integerConstantExpression';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { ModuleNode, ProcedureNode, Span } from '../../parser/nodes';
import { isLeafStatement } from '../../parser/nodes';
import { statementLabelDeclaration } from '../../flow/procedureLabels';
import type { PushFn } from '../analysisContext';
import { counterText, loopCountersAt, numericCounterPasses, type LoopCounter } from '../loopCounters';
import { stringLiteralValue, normalizeType } from '../typeInference';
import {
	activeModuleMembers,
	forEachVariableGroup,
	matchParenFrom,
	setAssignmentTarget,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
} from '../walker';

interface CollectionContents {
	/** Keys in element order; undefined for an element added without a key. */
	items: (string | undefined)[];
	/** False once an Add named a key the rule could not read, or ordered by Before/After. */
	keysKnown: boolean;
}

export function checkCollectionState(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const autoInstanced = collectionLocals(member, activity);
		const states = new Map<string, CollectionContents>();
		for (const name of autoInstanced.newLocals) {
			states.set(name, { items: [], keysKnown: true });
		}
		if (states.size === 0 && autoInstanced.plainLocals.size === 0) {
			continue;
		}
		for (const node of member.body) {
			if (activity?.isInactive(node.span)) {
				continue;
			}
			if (node.kind === 'VariableGroup') {
				continue; // a Dim inside the body declares, and runs nothing
			}
			if (!isLeafStatement(node)) {
				states.clear();
				continue;
			}
			if (node.kind === 'Statement' && node.singleLineIfBranches) {
				forgetMentioned(source, node.span, states);
				continue;
			}
			const toks = statementTokensAfterLeadingLabel(source, node.span);
			if (toks.length === 0) {
				continue;
			}
			// A label may be reached from anywhere; a GoSub may run any statement.
			if (statementLabelDeclaration(source, node.span) || tokenText(toks[0]) === 'gosub') {
				states.clear();
			}
			// `Set c = New Collection` starts an empty collection. `Set o = c`
			// makes o and c one collection, so they share one state and an Add
			// through either is seen by both (issue #147). Any other Set ends
			// tracking of its target, and of every tracked collection its value
			// names, since the value's new holder can change it unseen.
			const set = setAssignmentTarget(source, node.span);
			if (set) {
				const lower = set.name.toLowerCase();
				const value = toks.slice(toks.findIndex((tok) => tok.rawText === '=') + 1);
				const isCollectionLocal = autoInstanced.plainLocals.has(lower) || autoInstanced.newLocals.has(lower);
				const aliased = value.length === 1 ? tokenName(value[0])?.toLowerCase() : undefined;
				if (isCollectionLocal && value.length === 2 && tokenText(value[0]) === 'new' && tokenText(value[1]) === 'collection') {
					states.set(lower, { items: [], keysKnown: true });
					continue;
				}
				if (isCollectionLocal && aliased !== undefined && states.has(aliased)) {
					states.set(lower, states.get(aliased)!);
					continue;
				}
				states.delete(lower);
				for (const tok of value) {
					const mentioned = tokenName(tok)?.toLowerCase();
					if (mentioned && states.has(mentioned)) {
						states.delete(mentioned);
					}
				}
				continue;
			}
			checkStatement(node.span, toks, states, push);
		}
	}
}

function collectionLocals(
	proc: ProcedureNode,
	activity: ConditionalActivityTracker | undefined,
): { newLocals: Set<string>; plainLocals: Set<string> } {
	const newLocals = new Set<string>();
	const plainLocals = new Set<string>();
	forEachVariableGroup(proc.body, (group) => {
		if (group.isConst || group.modifier.toLowerCase() === 'static') {
			return;
		}
		for (const decl of group.declarations) {
			if (decl.isArray || normalizeType(decl.asType) !== 'collection') {
				continue;
			}
			(decl.isNew ? newLocals : plainLocals).add(decl.name.toLowerCase());
		}
	}, activity);
	return { newLocals, plainLocals };
}

/** Drops every tracked collection a statement names anywhere. */
function forgetMentioned(source: string, span: Span, states: Map<string, CollectionContents>): void {
	for (const tok of statementTokensAfterLeadingLabel(source, span)) {
		const lower = tokenName(tok)?.toLowerCase();
		if (lower && states.has(lower)) {
			states.delete(lower);
		}
	}
}

function checkStatement(base: Span, toks: readonly VbaToken[], states: Map<string, CollectionContents>, push: PushFn): void {
	const at = (from: number, to: number): Span => ({ start: base.start + toks[from].start, end: base.start + toks[to].end });
	// First pass: reads and the recognised forms, in source order. A mention
	// in any other shape ends tracking of that variable after this statement.
	const toForget = new Set<string>();
	const mutations: Array<() => void> = [];
	let i = 0;
	if (tokenText(toks[0]) === 'call') {
		i = 1;
	}
	for (; i < toks.length; i++) {
		const lower = tokenName(toks[i])?.toLowerCase();
		if (!lower || !states.has(lower) || toks[i - 1]?.rawText === '.') {
			continue;
		}
		const state = states.get(lower)!;
		const next = toks[i + 1];
		// `c(index)` or `c("key")`
		if (next?.rawText === '(') {
			const close = matchParenFrom(toks, i + 1);
			if (close > i + 2 && checkRead(lower, state, toks.slice(i + 2, close), at(i + 2, close - 1), push)) {
				continue;
			}
			toForget.add(lower);
			continue;
		}
		if (next?.rawText !== '.') {
			toForget.add(lower);
			continue;
		}
		const memberName = tokenText(toks[i + 2]);
		if (memberName === 'count') {
			continue;
		}
		if (memberName === 'item') {
			const itemClose = toks[i + 3]?.rawText === '(' ? matchParenFrom(toks, i + 3) : -1;
			if (itemClose > i + 4 && checkRead(lower, state, toks.slice(i + 4, itemClose), at(i + 4, itemClose - 1), push)) {
				continue;
			}
			toForget.add(lower);
			continue;
		}
		if ((memberName === 'add' || memberName === 'remove') && i === (tokenText(toks[0]) === 'call' ? 1 : 0)) {
			const args = argumentsAfter(toks, i + 3);
			if (memberName === 'add') {
				mutations.push(() => add(lower, state, args, base, push));
			} else if (args.length === 1) {
				mutations.push(() => remove(lower, state, args[0], base, push));
			} else {
				toForget.add(lower);
			}
			continue;
		}
		toForget.add(lower);
	}
	for (const mutation of mutations) {
		mutation();
	}
	for (const lower of toForget) {
		states.delete(lower);
	}
}

/** The argument groups of a statement-form or parenthesised member call starting at `from`. */
function argumentsAfter(toks: readonly VbaToken[], from: number): VbaToken[][] {
	let body = toks.slice(from);
	if (body[0]?.rawText === '(' && matchParenFrom(toks, from) === toks.length - 1) {
		body = body.slice(1, -1);
	}
	const out: VbaToken[][] = [];
	let current: VbaToken[] = [];
	let depth = 0;
	for (const tok of body) {
		if (tok.rawText === '(') {
			depth++;
		} else if (tok.rawText === ')') {
			depth--;
		}
		if (tok.rawText === ',' && depth === 0) {
			out.push(current);
			current = [];
			continue;
		}
		current.push(tok);
	}
	if (current.length > 0 || out.length > 0) {
		out.push(current);
	}
	return out;
}

function literalKey(arg: readonly VbaToken[]): string | undefined {
	return arg.length === 1 && arg[0].kind === 'stringLiteral' ? stringLiteralValue(arg[0].rawText).toLowerCase() : undefined;
}

/** The whole-number value of a literal argument, optionally negated. */
function literalIndex(arg: readonly VbaToken[]): number | undefined {
	if (arg.length === 1 && arg[0].kind === 'integerLiteral') {
		return parseVbaIntegerLiteral(arg[0].rawText);
	}
	if (arg.length === 2 && arg[0].rawText === '-' && arg[1].kind === 'integerLiteral') {
		const value = parseVbaIntegerLiteral(arg[1].rawText);
		return value === undefined ? undefined : -value;
	}
	return undefined;
}

/** Judges `c(arg)` or `c.Item(arg)`; false when the argument is not a literal the rule reads. */
function checkRead(name: string, state: CollectionContents, arg: readonly VbaToken[], span: Span, push: PushFn): boolean {
	const index = literalIndex(arg);
	if (index !== undefined) {
		reportIndex(name, state, index, span, push);
		return true;
	}
	if (arg.length === 1 && arg[0].kind === 'stringLiteral') {
		reportKey(name, state, stringLiteralValue(arg[0].rawText), span, push);
		return true;
	}
	return false;
}

function reportIndex(name: string, state: CollectionContents, index: number, span: Span, push: PushFn): void {
	if (state.items.length === 0) {
		push('collectionIndexOutOfRange', `'${name}' holds nothing here, so no index reaches an element. This will raise Run-time error '5': Invalid procedure call or argument.`, span);
		return;
	}
	if (index < 1 || index > state.items.length) {
		push('collectionIndexOutOfRange', `'${name}' holds ${state.items.length} element${state.items.length === 1 ? '' : 's'} here, indexed 1 to ${state.items.length}; ${index} is outside that. This will raise Run-time error '9': Subscript out of range.`, span);
	}
}

function reportKey(name: string, state: CollectionContents, key: string, span: Span, push: PushFn): boolean {
	if (state.items.length === 0) {
		push('collectionKeyNotFound', `'${name}' holds nothing here, so no key reaches an element. This will raise Run-time error '5': Invalid procedure call or argument.`, span);
		return true;
	}
	if (state.keysKnown && !state.items.includes(key.toLowerCase())) {
		push('collectionKeyNotFound', `No element of '${name}' was added with the key "${key}". This will raise Run-time error '5': Invalid procedure call or argument.`, span);
		return true;
	}
	return false;
}

function add(name: string, state: CollectionContents, args: VbaToken[][], base: Span, push: PushFn): void {
	const keyArg = args[1];
	const key = keyArg && keyArg.length > 0 ? literalKey(keyArg) : undefined;
	if (keyArg && keyArg.length > 0 && key === undefined) {
		state.keysKnown = false;
	}
	if (key !== undefined && state.keysKnown && state.items.includes(key)) {
		push('collectionKeyInUse', `'${name}' already has an element with the key "${stringLiteralValue(keyArg[0].rawText)}" (keys compare without case). This will raise Run-time error '457': This key is already associated with an element of this collection.`, { start: base.start + keyArg[0].start, end: base.start + keyArg[keyArg.length - 1].end });
		return;
	}
	if (args.length > 2 && args.slice(2).some((arg) => arg.length > 0)) {
		// Before or After: the position is not followed, the count is.
		state.items.push(key);
		state.keysKnown = false;
		return;
	}
	state.items.push(key);
}

function remove(name: string, state: CollectionContents, arg: VbaToken[], base: Span, push: PushFn): void {
	const span = { start: base.start + arg[0].start, end: base.start + arg[arg.length - 1].end };
	const index = literalIndex(arg);
	if (index !== undefined) {
		if (state.items.length === 0 || index < 1 || index > state.items.length) {
			reportIndex(name, state, index, span, push);
			return;
		}
		state.items.splice(index - 1, 1);
		return;
	}
	const key = literalKey(arg);
	if (key === undefined) {
		// A variable index or key: one element fewer, which one unknown.
		state.items.pop();
		state.keysKnown = false;
		return;
	}
	if (reportKey(name, state, stringLiteralValue(arg[0].rawText), span, push)) {
		return;
	}
	const position = state.items.indexOf(key);
	if (position >= 0) {
		state.items.splice(position, 1);
	} else {
		state.items.pop();
		state.keysKnown = false;
	}
}

/**
 * A loop counter indexing a Collection outside 1 to Count (issue #200,
 * measured in Excel 16.0): `For i = 0 To c.Count - 1` reads c(0) on its
 * first pass, and `For i = 1 To c.Count + 1` reads past the last element on
 * its last. Collections are 1-based; an index outside raises 9, or 5 when
 * the collection is empty. The contents are not tracked into a loop, so the
 * error is 9 only where the loop's own bounds show it runs with an element.
 */
export function checkCollectionLoopCounters(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const locals = collectionLocals(member, activity);
		const collections = new Set([...locals.newLocals, ...locals.plainLocals]);
		for (const param of member.params) {
			if (!param.isArray && normalizeType(param.asType) === 'collection') {
				collections.add(param.name.toLowerCase());
			}
		}
		if (collections.size === 0) {
			continue;
		}
		for (const [stmt, counters] of loopCountersAt(source, member.body, activity)) {
			const toks = statementTokensAfterLeadingLabel(source, stmt.span);
			for (let i = 0; i + 1 < toks.length; i++) {
				const lower = tokenName(toks[i])?.toLowerCase();
				if (!lower || !collections.has(lower) || toks[i - 1]?.rawText === '.') {
					continue;
				}
				// `c(i)` or `c.Item(i)`
				const open = toks[i + 1].rawText === '(' ? i + 1
					: toks[i + 1].rawText === '.' && tokenText(toks[i + 2]) === 'item' && toks[i + 3]?.rawText === '(' ? i + 3 : -1;
				const close = open < 0 ? -1 : matchParenFrom(toks, open);
				const arg = close === open + 2 ? toks[open + 1] : undefined;
				const counter = arg ? counters.get(tokenName(arg)?.toLowerCase() ?? '') : undefined;
				const message = counter && arg ? collectionCounterMessage(toks[i].rawText, lower, arg.rawText, counter) : undefined;
				if (arg && message) {
					push('collectionIndexOutOfRange', message, { start: stmt.span.start + arg.start, end: stmt.span.start + arg.end });
				}
			}
		}
	}
}

function collectionCounterMessage(name: string, lower: string, counterName: string, counter: LoopCounter): string | undefined {
	const count = (atom: { kind: string; name: string }): boolean => atom.kind === 'count' && atom.name === lower;
	let reached: string | undefined;
	for (const pass of numericCounterPasses(counter, () => undefined)) {
		if (pass.value < 1) {
			reached = pass.pass === 'first'
				? `Counter '${counterName}' is ${pass.value} on its first pass`
				: `Counter '${counterName}' reaches ${pass.value} on its last pass`;
			break;
		}
	}
	if (!reached) {
		const passes: Array<['first' | 'last', typeof counter.first | undefined]> = [['first', counter.first], ['last', counter.last]];
		for (const [pass, value] of passes) {
			if (value?.atom && count(value.atom) && value.offset > 0) {
				reached = pass === 'first'
					? `Counter '${counterName}' is ${counterText(value)} on its first pass`
					: `Counter '${counterName}' reaches ${counterText(value)} on its last pass`;
				break;
			}
		}
	}
	if (!reached) {
		return undefined;
	}
	const error = runsWithAnElement(counter, count)
		? `This will raise Run-time error '9': Subscript out of range.`
		: `This will raise Run-time error '9': Subscript out of range, or '5' if '${name}' is empty.`;
	return `${reached}, and '${name}' holds its elements at 1 to ${name}.Count. ${error}`;
}

/** Whether the loop's bounds show it runs only when the collection has an element: `For i = 0 To c.Count - 1`. */
function runsWithAnElement(counter: LoopCounter, count: (atom: { kind: string; name: string }) => boolean): boolean {
	const [low, high] = counter.step > 0 ? [counter.first, counter.last] : [counter.last, counter.first];
	// The loop runs when low <= Count + offset, so Count >= low - offset.
	return !low?.atom && high?.atom !== undefined && count(high.atom) && low !== undefined && low.offset - high.offset >= 1;
}
