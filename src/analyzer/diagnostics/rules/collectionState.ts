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
//  - collection-add-argument (issue #219): a key that is a number or True,
//    `c.Add "x", 5` -> 13, and so is a Variant never assigned, which is
//    Empty; Before and After together -> 5. Before or After
//    on an empty collection -> 5, and outside 1 to Count -> 9, are
//    collection-index-out-of-range's.
//  - array-subscript-out-of-bounds (issue #248): `c.Add Array(1, 2)` then
//    `c(1)(5)` indexes past the end of the array the item holds -> 9.
//
// The rule follows a procedure's top-level statements in order, as the file
// rule does: a block ends what is known, and any use of the variable other
// than Add, Remove, Item, Count and indexing ends it too.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import { evaluateIntegerConstantExpression, parseVbaIntegerLiteral, resolveRawIntegerConstants } from '../../constants/integerConstantExpression';
import type { HostObjectModel } from '../../host/excelObjectModel';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { VbaSymbol } from '../../symbols/symbolModel';
import { collectModuleLiteralIntegerConstants } from '../constExpr';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { IntegerConstantLookup } from '../../constants/integerConstantExpression';
import type { BodyNode, ModuleNode, ProcedureNode, Span } from '../../parser/nodes';
import { walkEnteringBlocks } from '../dataflow';
import { isLeafStatement } from '../../parser/nodes';
import { statementLabelDeclaration } from '../../flow/procedureLabels';
import { procedureSymbolFor, type PushFn } from '../analysisContext';
import { counterText, loopCountersAt, numericCounterPasses, type LoopCounter } from '../loopCounters';
import { isKnownScalarType, knownLocalLiteralValuesAt, normalizeType, procedureIntegerConstantLookup, stringLiteralValue, unreachableStatementsIn, withKnownLocals } from '../typeInference';
import {
	activeModuleMembers,
	blockHeaderLineSpan,
	forEachVariableGroup,
	matchParenFrom,
	setAssignmentTarget,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
} from '../walker';
import { nameMentions, namesIn } from './shared';
import { arrayValueShape, moduleOptionBase, shapeSubscriptViolation, type FixedArrayBound } from './arrays';

interface CollectionContents {
	/** Keys in element order; undefined for an element added without a key. */
	items: (string | undefined)[];
	/** False once an Add named a key the rule could not read, or ordered by Before/After. */
	keysKnown: boolean;
	/** The bounds of the array each element holds, where an Add gave it `Array(...)`; aligned with items. */
	shapes: (FixedArrayBound | undefined)[];
	/**
	 * What each element is, where an Add gave it a tracked Collection, which
	 * the element then shares, or a number literal (issue #452); aligned with
	 * items.
	 */
	held: Held[];
	/** Set once a name for this collection stopped being followed: it may have changed unseen. */
	stale?: boolean;
}

type Held = CollectionContents | 'number' | 'string' | undefined;

function emptyContents(): CollectionContents {
	return { items: [], keysKnown: true, shapes: [], held: [] };
}

/** Stops following a name; what it named may now change unseen, so an element sharing it is no longer read. */
function forgetCollection(states: Map<string, CollectionContents>, lower: string): void {
	const contents = states.get(lower);
	if (contents) {
		contents.stale = true;
	}
	states.delete(lower);
}

export function checkCollectionState(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
	symbols?: ReturnType<typeof buildModuleSymbols>,
	projectIntegerConstants?: ReadonlyMap<string, string | undefined>,
	projectVisibleSymbols?: readonly VbaSymbol[],
	hostModel?: HostObjectModel,
): void {
	const moduleConstants = collectModuleLiteralIntegerConstants(mod, activity, resolveRawIntegerConstants(projectIntegerConstants ?? new Map(), new Map()));
	const optionBase = moduleOptionBase(mod, activity);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const autoInstanced = collectionLocals(member, activity);
		const states = new Map<string, CollectionContents>();
		for (const name of autoInstanced.newLocals) {
			states.set(name, emptyContents());
		}
		if (states.size === 0 && autoInstanced.plainLocals.size === 0) {
			continue;
		}
		// A Variant local named nowhere but one statement is Empty there.
		let mentions: Map<string, number> | undefined;
		const isEmpty = (lower: string): boolean => autoInstanced.variantLocals.has(lower)
			&& (mentions ??= nameMentions(source, member, activity)).get(lower) === 1;
		// An index through a Const or a local with one known value (issue #238).
		const constants = symbols ? procedureIntegerConstantLookup(member, moduleConstants, symbols, projectVisibleSymbols, activity, hostModel) : undefined;
		const valuesAt = symbols ? knownLocalLiteralValuesAt(source, member, symbols, activity) : undefined;
		// Code that never runs changes nothing and raises nothing (issue #406).
		const unreachable = symbols ? unreachableStatementsIn(source, member, symbols, activity) : undefined;
		// A scalar or Variant local, or the Function's result, which a
		// Collection's value is Let into (issue #452).
		const holdsValue = (type: string | undefined): boolean => {
			const normalized = normalizeType(type);
			return normalized === undefined || normalized === 'variant' || isKnownScalarType(normalized);
		};
		const scalars = new Set((symbols ? procedureSymbolFor(symbols, member)?.children ?? [] : [])
			.filter((child) => child.kind === 'localVariable' && !child.isArray && holdsValue(child.asType))
			.map((child) => child.name.toLowerCase()));
		if (member.procKind === 'Function' && holdsValue(member.returnType)) {
			scalars.add(member.name.toLowerCase());
		}
		const scalarLocal = (lower: string): boolean => scalars.has(lower);
		// Blocks are entered with the state they start with (issue #237).
		const visit = (node: BodyNode): void => {
			if (!isLeafStatement(node) || unreachable?.has(node)) {
				return; // a Dim inside the body declares, and runs nothing
			}
			if (node.kind === 'Statement' && node.singleLineIfBranches) {
				forgetMentioned(source, node.span, states);
				return;
			}
			const toks = statementTokensAfterLeadingLabel(source, node.span);
			if (toks.length === 0) {
				return;
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
					states.set(lower, emptyContents());
					return;
				}
				if (isCollectionLocal && aliased !== undefined && states.has(aliased)) {
					states.set(lower, states.get(aliased)!);
					return;
				}
				forgetCollection(states, lower);
				for (const tok of value) {
					const mentioned = tokenName(tok)?.toLowerCase();
					if (mentioned && states.has(mentioned)) {
						forgetCollection(states, mentioned);
					}
				}
				return;
			}
			const lookup = constants && valuesAt ? withKnownLocals(constants, valuesAt(node)) : undefined;
			const indexOf = (arg: readonly VbaToken[]): number | undefined => literalIndex(arg)
				?? (lookup ? evaluateIntegerConstantExpression(arg.map((tok) => tok.rawText).join(' '), lookup) : undefined);
			checkStatement(node.span, toks, states, push, isEmpty, indexOf, optionBase, lookup, scalarLocal);
		};
		walkEnteringBlocks(source, member.body, (node) => activity?.isInactive(node.span) === true, visit, {
			snapshot: () => cloneStates(states),
			restore: (saved) => {
				states.clear();
				for (const [lower, contents] of cloneStates(saved)) {
					states.set(lower, contents);
				}
			},
			forget: (names) => {
				for (const lower of names) {
					forgetCollection(states, lower);
				}
			},
			touches: (stmt) => namesIn(source, stmt.span),
			// A counted loop that removes or reads by its counter (issue #263).
			enter: (node) => simulateCountedLoop(source, node, states, push, activity),
		});
	}
}

/** The most passes a counted loop is run for. */
const MAX_SIMULATED_PASSES = 10000;

/**
 * `For i = 1 To c.Count: c.Remove i: Next` on three elements removes 1 and 2,
 * then finds no element 3 (issue #263, measured in Excel 16.0: error 9; and
 * error 5 once the collection is empty). A For loop whose bounds the
 * contents decide, and whose body is plain statements that touch the
 * collection only by `c.Remove k` and `c(k)`, k the counter, a whole number
 * or the counter plus or minus one, is run pass by pass. Anything else in
 * the body that names the collection, writes the counter or may leave the
 * pass stops it.
 */
function simulateCountedLoop(
	source: string,
	node: BodyNode,
	states: ReadonlyMap<string, CollectionContents>,
	push: PushFn,
	activity: ConditionalActivityTracker | undefined,
): void {
	if (node.kind !== 'ForBlock' || node.each || !node.controlVariable || states.size === 0) {
		return;
	}
	const counter = node.controlVariable.toLowerCase();
	const header = statementTokensAfterLeadingLabel(source, blockHeaderLineSpan(source, node.span)).filter((tok) => tok.kind !== 'comment');
	const eq = header.findIndex((tok) => tok.rawText === '=');
	const to = header.findIndex((tok) => tokenText(tok) === 'to');
	const stepAt = header.findIndex((tok) => tokenText(tok) === 'step');
	const bound = (toks: readonly VbaToken[]): number | undefined => {
		const literal = literalIndex(toks);
		if (literal !== undefined) {
			return literal;
		}
		// `c.Count`, `c.Count - 1`
		const name = tokenName(toks[0])?.toLowerCase();
		const contents = name ? states.get(name) : undefined;
		if (!contents || toks[1]?.rawText !== '.' || tokenText(toks[2]) !== 'count') {
			return undefined;
		}
		if (toks.length === 3) {
			return contents.items.length;
		}
		const offset = toks.length === 5 && (toks[3].rawText === '+' || toks[3].rawText === '-') ? literalIndex([toks[4]]) : undefined;
		return offset === undefined ? undefined : contents.items.length + (toks[3].rawText === '-' ? -offset : offset);
	};
	const start = eq > 0 && to > eq ? bound(header.slice(eq + 1, to)) : undefined;
	const limit = to > 0 ? bound(header.slice(to + 1, stepAt > 0 ? stepAt : header.length)) : undefined;
	const step = stepAt > 0 ? literalIndex(header.slice(stepAt + 1)) : 1;
	if (start === undefined || limit === undefined || step === undefined || step === 0) {
		return;
	}
	interface Use { name: string; display: string; arg: readonly VbaToken[]; base: number; removes: boolean }
	const uses: Use[] = [];
	for (const stmt of node.body) {
		if (activity?.isInactive(stmt.span)) {
			continue;
		}
		if (!isLeafStatement(stmt) || (stmt.kind === 'Statement' && stmt.singleLineIfBranches)) {
			return;
		}
		const toks = statementTokensAfterLeadingLabel(source, stmt.span).filter((tok) => tok.kind !== 'comment');
		const head = tokenText(toks[0]);
		if (['exit', 'goto', 'gosub', 'resume', 'return', 'end', 'on', 'stop'].includes(head) || statementLabelDeclaration(source, stmt.span)) {
			return;
		}
		// The counter only read: an operand, a whole collection index, or the
		// whole Remove argument. Assigned, passed or printed, it is not followed.
		for (let i = 0; i < toks.length; i++) {
			if (tokenName(toks[i])?.toLowerCase() !== counter || toks[i - 1]?.rawText === '.') {
				continue;
			}
			const operand = (i > 0 && toks[i - 1].kind === 'operator') || (toks[i + 1]?.kind === 'operator' && !(i === 0 && toks[i + 1].rawText === '='));
			const indexes = toks[i - 1]?.rawText === '(' && toks[i + 1]?.rawText === ')' && states.has(tokenName(toks[i - 2])?.toLowerCase() ?? '');
			const removes = i === 3 && toks.length === 4 && tokenText(toks[2]) === 'remove' && states.has(tokenName(toks[0])?.toLowerCase() ?? '');
			if (!operand && !indexes && !removes) {
				return;
			}
		}
		for (let i = 0; i < toks.length; i++) {
			const lower = tokenName(toks[i])?.toLowerCase();
			if (!lower || toks[i - 1]?.rawText === '.' || !states.has(lower)) {
				continue;
			}
			if (i === 0 && toks[1]?.rawText === '.' && tokenText(toks[2]) === 'remove' && toks.length > 3) {
				uses.push({ name: lower, display: toks[0].rawText, arg: toks.slice(3), base: stmt.span.start, removes: true });
				break;
			}
			const open = toks[i + 1]?.rawText === '(' ? i + 1 : toks[i + 1]?.rawText === '.' && tokenText(toks[i + 2]) === 'item' && toks[i + 3]?.rawText === '(' ? i + 3 : -1;
			if (open < 0) {
				return; // Add, Count after a change, a pass or a Set: not followed
			}
			const close = matchParenFrom(toks, open);
			uses.push({ name: lower, display: toks[i].rawText, arg: toks.slice(open + 1, close), base: stmt.span.start, removes: false });
			i = close;
		}
	}
	if (uses.length === 0) {
		return;
	}
	const indexAt = (arg: readonly VbaToken[], value: number): number | undefined => {
		const literal = literalIndex(arg);
		if (literal !== undefined) {
			return literal;
		}
		if (tokenName(arg[0])?.toLowerCase() !== counter) {
			return undefined;
		}
		if (arg.length === 1) {
			return value;
		}
		const offset = arg.length === 3 && (arg[1].rawText === '+' || arg[1].rawText === '-') ? literalIndex([arg[2]]) : undefined;
		return offset === undefined ? undefined : value + (arg[1].rawText === '-' ? -offset : offset);
	};
	if (uses.some((use) => indexAt(use.arg, start) === undefined)) {
		return;
	}
	const counts = new Map([...new Set(uses.map((use) => use.name))].map((name) => [name, states.get(name)!.items.length]));
	let passes = 0;
	for (let value = start; step > 0 ? value <= limit : value >= limit; value += step) {
		if (++passes > MAX_SIMULATED_PASSES) {
			return;
		}
		for (const use of uses) {
			const count = counts.get(use.name)!;
			const index = indexAt(use.arg, value)!;
			if (index >= 1 && index <= count) {
				if (use.removes) {
					counts.set(use.name, count - 1);
				}
				continue;
			}
			if (passes === 1 && literalIndex(use.arg) !== undefined) {
				return; // the walk into the block reports the first pass
			}
			const span = { start: use.base + use.arg[0].start, end: use.base + use.arg[use.arg.length - 1].end };
			const where = `On the pass of the For loop where '${node.controlVariable}' is ${value}`;
			const message = count === 0
				? `${where}, '${use.display}' holds nothing, so no index reaches an element. This will raise Run-time error '5': Invalid procedure call or argument.`
				: `${where}, '${use.display}' holds ${count} element${count === 1 ? '' : 's'}, indexed 1 to ${count}; ${index} is outside that. This will raise Run-time error '9': Subscript out of range.`;
			push('collectionIndexOutOfRange', message, span);
			return;
		}
	}
}

/** A copy of the states in which two names that shared one collection still do. */
function cloneStates(states: ReadonlyMap<string, CollectionContents>): Map<string, CollectionContents> {
	const copies = new Map<CollectionContents, CollectionContents>();
	// An element that is a collection is copied once, as the name sharing it is.
	const copyOf = (contents: CollectionContents): CollectionContents => {
		let copy = copies.get(contents);
		if (!copy) {
			copy = { items: [...contents.items], keysKnown: contents.keysKnown, shapes: [...contents.shapes], held: [], ...(contents.stale ? { stale: true } : {}) };
			copies.set(contents, copy);
			copy.held = contents.held.map((held) => (typeof held === 'object' ? copyOf(held) : held));
		}
		return copy;
	};
	const out = new Map<string, CollectionContents>();
	for (const [lower, contents] of states) {
		out.set(lower, copyOf(contents));
	}
	return out;
}

function collectionLocals(
	proc: ProcedureNode,
	activity: ConditionalActivityTracker | undefined,
): { newLocals: Set<string>; plainLocals: Set<string>; variantLocals: Set<string> } {
	const newLocals = new Set<string>();
	const plainLocals = new Set<string>();
	const variantLocals = new Set<string>();
	forEachVariableGroup(proc.body, (group) => {
		if (group.isConst || group.modifier.toLowerCase() === 'static') {
			return;
		}
		for (const decl of group.declarations) {
			const type = normalizeType(decl.asType);
			if (!decl.isArray && (type === undefined || type === 'variant') && !/[%&^!#@$]$/.test(decl.name)) {
				variantLocals.add(decl.name.toLowerCase());
			}
			if (decl.isArray || type !== 'collection') {
				continue;
			}
			(decl.isNew ? newLocals : plainLocals).add(decl.name.toLowerCase());
		}
	}, activity);
	return { newLocals, plainLocals, variantLocals };
}

/** Drops every tracked collection a statement names anywhere. */
function forgetMentioned(source: string, span: Span, states: Map<string, CollectionContents>): void {
	for (const tok of statementTokensAfterLeadingLabel(source, span)) {
		const lower = tokenName(tok)?.toLowerCase();
		if (lower && states.has(lower)) {
			forgetCollection(states, lower);
		}
	}
}

type IndexOf = (arg: readonly VbaToken[]) => number | undefined;

function checkStatement(
	base: Span,
	toks: readonly VbaToken[],
	states: Map<string, CollectionContents>,
	push: PushFn,
	isEmpty: (lower: string) => boolean,
	indexOf: IndexOf = literalIndex,
	optionBase = 0,
	lookup?: IntegerConstantLookup,
	scalarLocal: (lower: string) => boolean = () => false,
): void {
	const at = (from: number, to: number): Span => ({ start: base.start + toks[from].start, end: base.start + toks[to].end });
	// First pass: reads and the recognised forms, in source order. A mention
	// in any other shape ends tracking of that variable after this statement.
	const toForget = new Set<string>();
	const mutations: Array<() => void> = [];
	let i = 0;
	if (tokenText(toks[0]) === 'call') {
		i = 1;
	}
	const first = i;
	// `c.Add inner`: the element shares inner, which the Add leaves as it is.
	const heldOf = (item: readonly VbaToken[]): Held => {
		const named = item.length === 1 ? tokenName(item[0])?.toLowerCase() : undefined;
		const contents = named ? states.get(named) : undefined;
		if (contents) {
			return contents;
		}
		return literalIndex(item) !== undefined || (item.length === 1 && item[0].kind === 'floatLiteral') ? 'number' : item.length === 1 && item[0].kind === 'stringLiteral' ? 'string' : undefined;
	};
	const addsWhole = states.has(tokenName(toks[first])?.toLowerCase() ?? '') && toks[first + 1]?.rawText === '.' && tokenText(toks[first + 2]) === 'add'
		? argumentsAfter(toks, first + 3)[0] : undefined;
	const addedName = addsWhole?.length === 1 && tokenName(addsWhole[0]) ? addsWhole[0] : undefined;
	const item = { base, toks, push, indexOf, mutations, isEmpty, optionBase, heldOf, scalarLocal, first };
	for (; i < toks.length; i++) {
		const lower = tokenName(toks[i])?.toLowerCase();
		if (!lower || !states.has(lower) || toks[i - 1]?.rawText === '.' || toks[i] === addedName) {
			continue;
		}
		const state = states.get(lower)!;
		const next = toks[i + 1];
		// `c(index)` or `c("key")`
		if (next?.rawText === '(') {
			const close = matchParenFrom(toks, i + 1);
			if (close > i + 2 && checkRead(lower, state, toks.slice(i + 2, close), at(i + 2, close - 1), push, indexOf)) {
				checkItemArray(base, toks, toks[i].rawText, state, toks.slice(i + 2, close), close, push, indexOf, lookup);
				useHeldItem(item, i, toks[i].rawText, state, toks.slice(i + 2, close), close);
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
			if (itemClose > i + 4 && checkRead(lower, state, toks.slice(i + 4, itemClose), at(i + 4, itemClose - 1), push, indexOf)) {
				checkItemArray(base, toks, toks[i].rawText, state, toks.slice(i + 4, itemClose), itemClose, push, indexOf, lookup);
				useHeldItem(item, i, `${toks[i].rawText}.Item`, state, toks.slice(i + 4, itemClose), itemClose);
				continue;
			}
			toForget.add(lower);
			continue;
		}
		if ((memberName === 'add' || memberName === 'remove') && i === (tokenText(toks[0]) === 'call' ? 1 : 0)) {
			const args = argumentsAfter(toks, i + 3);
			if (memberName === 'add') {
				mutations.push(() => add(lower, state, args, base, push, isEmpty, indexOf, optionBase, heldOf));
			} else if (args.length === 1) {
				mutations.push(() => remove(lower, state, args[0], base, push, indexOf));
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
		forgetCollection(states, lower);
	}
}

interface ItemContext {
	base: Span;
	toks: readonly VbaToken[];
	push: PushFn;
	indexOf: IndexOf;
	mutations: Array<() => void>;
	isEmpty: (lower: string) => boolean;
	optionBase: number;
	heldOf: (item: readonly VbaToken[]) => Held;
	scalarLocal: (lower: string) => boolean;
	/** The statement's first token after a Call. */
	first: number;
}

/**
 * What follows `c(k)` where element k is a Collection the code added, or a
 * number (issue #452, measured in Excel 16.0): an index or key into the
 * inner Collection, and its Add and Remove, are judged against what it
 * holds; a number indexed raises 13 and a member of one 424; a Collection
 * Let into a typed local raises 450, its default member Item needing an
 * index. Any other use of an inner Collection may change it unseen.
 */
function useHeldItem(ctx: ItemContext, nameAt: number, display: string, state: CollectionContents, arg: readonly VbaToken[], close: number): void {
	const { toks, base } = ctx;
	const key = literalKey(arg);
	const index = ctx.indexOf(arg) ?? (key !== undefined && state.keysKnown && state.items.includes(key) ? state.items.indexOf(key) + 1 : undefined);
	const held = index !== undefined && index >= 1 ? state.held[index - 1] : undefined;
	if (held === undefined || (typeof held === 'object' && held.stale)) {
		return;
	}
	const shown = `${display}(${arg.map((tok) => tok.rawText).join('')})`;
	const spanOf = (from: number, to: number): Span => ({ start: base.start + toks[from].start, end: base.start + toks[to].end });
	const after = toks[close + 1]?.rawText;
	const member = after === '.' ? tokenText(toks[close + 2]) : undefined;
	if (held === 'number' || held === 'string') {
		if (after === '(') {
			ctx.push('variantValueMisuse', `'${shown}' holds a ${held}, which takes no index. This will raise Run-time error '13': Type mismatch.`, spanOf(nameAt, matchParenFrom(toks, close + 1)));
		} else if (member) {
			ctx.push('variantValueMisuse', `'${shown}' holds a ${held}, not an object, so it has no ${toks[close + 2].rawText}. This will raise Run-time error '424': Object required.`, spanOf(nameAt, close + 2));
		}
		return;
	}
	if (after === '(' || (member === 'item' && toks[close + 3]?.rawText === '(')) {
		const open = after === '(' ? close + 1 : close + 3;
		const innerClose = matchParenFrom(toks, open);
		if (innerClose > open + 1) {
			checkRead(shown, held, toks.slice(open + 1, innerClose), spanOf(open + 1, innerClose - 1), ctx.push, ctx.indexOf);
		}
		return;
	}
	if (member === 'count') {
		return;
	}
	if ((member === 'add' || member === 'remove') && nameAt === ctx.first) {
		const args = argumentsAfter(toks, close + 3);
		if (member === 'add') {
			ctx.mutations.push(() => add(shown, held, args, base, ctx.push, ctx.isEmpty, ctx.indexOf, ctx.optionBase, ctx.heldOf));
		} else if (args.length === 1) {
			ctx.mutations.push(() => remove(shown, held, args[0], base, ctx.push, ctx.indexOf));
		} else {
			held.stale = true;
		}
		return;
	}
	if (member !== undefined) {
		held.stale = true;
		return;
	}
	// `v = c(1)` with v a Long or a Variant.
	const target = tokenName(toks[ctx.first])?.toLowerCase();
	if (target && toks[ctx.first + 1]?.rawText === '=' && nameAt === ctx.first + 2 && close === toks.length - 1 && ctx.scalarLocal(target)) {
		ctx.push('objectDefaultValue', `'${shown}' is a Collection: its default member Item needs an index, so it has no value for '${toks[ctx.first].rawText}' to take. This will raise Run-time error '450': Wrong number of arguments or invalid property assignment.`, spanOf(nameAt, close));
	}
}

/** `c(1)(5)`: the parentheses after an item read, against the array the item holds. */
function checkItemArray(
	base: Span,
	toks: readonly VbaToken[],
	name: string,
	state: CollectionContents,
	arg: readonly VbaToken[],
	close: number,
	push: PushFn,
	indexOf: IndexOf,
	lookup: IntegerConstantLookup | undefined,
): void {
	const index = toks[close + 1]?.rawText === '(' ? indexOf(arg) : undefined;
	const shape = index === undefined ? undefined : state.shapes[index - 1];
	const hit = shape && shapeSubscriptViolation(base, toks, { ...shape, name: `${name}(${index})` }, close + 1, lookup);
	if (hit) {
		push(hit.rule ?? 'arraySubscriptOutOfBounds', hit.message, hit.span);
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
function checkRead(name: string, state: CollectionContents, arg: readonly VbaToken[], span: Span, push: PushFn, indexOf: IndexOf): boolean {
	const index = indexOf(arg);
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

const ADD_PARAMETERS = ['item', 'key', 'before', 'after'];

/** Add's arguments by parameter, positional or named; undefined when a name is unknown. */
function addArguments(args: readonly VbaToken[][]): Map<string, VbaToken[]> | undefined {
	const out = new Map<string, VbaToken[]>();
	for (let k = 0; k < args.length; k++) {
		const arg = args[k];
		if (arg[1]?.rawText === ':=') {
			const param = tokenText(arg[0]);
			if (!ADD_PARAMETERS.includes(param)) {
				return undefined;
			}
			out.set(param, arg.slice(2));
		} else if (arg.length > 0 && k < ADD_PARAMETERS.length) {
			out.set(ADD_PARAMETERS[k], arg);
		}
	}
	return out;
}

/**
 * What Add refuses before it adds (issue #219, measured in Excel 16.0): a key
 * that is a number or True rather than a string raises 13; Before and After
 * together raise 5; either one on an empty collection raises 5; and an index
 * outside 1 to Count raises 9 (Before:=0, After:=2 with one element).
 */
function addRefusal(name: string, state: CollectionContents, byName: ReadonlyMap<string, VbaToken[]>, base: Span, isEmpty: (lower: string) => boolean, indexOf: IndexOf): { rule: 'collectionAddArgument' | 'collectionIndexOutOfRange'; message: string; span: Span } | undefined {
	const spanOf = (arg: readonly VbaToken[]): Span => ({ start: base.start + arg[0].start, end: base.start + arg[arg.length - 1].end });
	const key = byName.get('key');
	const keyLiteral = key ? key.filter((t) => t.kind !== 'comment') : [];
	const nonString = keyLiteral.length > 0 && (literalIndex(keyLiteral) !== undefined
		|| (keyLiteral.length === 1 && keyLiteral[0].kind === 'floatLiteral')
		|| (keyLiteral.length === 1 && (tokenText(keyLiteral[0]) === 'true' || tokenText(keyLiteral[0]) === 'false')));
	const emptyKey = keyLiteral.length === 1 && isEmpty(tokenName(keyLiteral[0])?.toLowerCase() ?? '');
	if (emptyKey) {
		return { rule: 'collectionAddArgument', message: `The key of '${name}.Add' is '${keyLiteral[0].rawText}', which is never assigned and so is Empty, not a string. This will raise Run-time error '13': Type mismatch.`, span: spanOf(keyLiteral) };
	}
	if (nonString) {
		return { rule: 'collectionAddArgument', message: `The key of '${name}.Add' is ${keyLiteral.map((t) => t.rawText).join('')}, not a string. This will raise Run-time error '13': Type mismatch.`, span: spanOf(keyLiteral) };
	}
	const before = byName.get('before');
	const after = byName.get('after');
	if (before && after) {
		return { rule: 'collectionAddArgument', message: `'${name}.Add' is given both Before and After. This will raise Run-time error '5': Invalid procedure call or argument.`, span: spanOf(after) };
	}
	const position = before ?? after;
	if (!position) {
		return undefined;
	}
	if (state.items.length === 0) {
		return { rule: 'collectionIndexOutOfRange', message: `'${name}' holds nothing here, so ${before ? 'Before' : 'After'} names no element. This will raise Run-time error '5': Invalid procedure call or argument.`, span: spanOf(position) };
	}
	const index = indexOf(position);
	if (index !== undefined && (index < 1 || index > state.items.length)) {
		return { rule: 'collectionIndexOutOfRange', message: `'${name}' holds ${state.items.length} element${state.items.length === 1 ? '' : 's'} here, indexed 1 to ${state.items.length}; ${before ? 'Before' : 'After'} is ${index}. This will raise Run-time error '9': Subscript out of range.`, span: spanOf(position) };
	}
	return undefined;
}

function add(name: string, state: CollectionContents, rawArgs: VbaToken[][], base: Span, push: PushFn, isEmpty: (lower: string) => boolean, indexOf: IndexOf, optionBase: number, heldOf: (item: readonly VbaToken[]) => Held = () => undefined): void {
	const byName = addArguments(rawArgs);
	if (!byName) {
		state.items.push(undefined);
		state.shapes.push(undefined);
		state.held.push(undefined);
		state.keysKnown = false;
		return;
	}
	const refusal = addRefusal(name, state, byName, base, isEmpty, indexOf);
	if (refusal) {
		push(refusal.rule, refusal.message, refusal.span);
		return;
	}
	const args = [byName.get('item') ?? [], byName.get('key') ?? [], byName.get('before') ?? [], byName.get('after') ?? []];
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
		state.shapes = state.items.map(() => undefined);
		state.held = state.items.map(() => undefined);
		return;
	}
	state.items.push(key);
	state.shapes.push(args[0].length > 0 ? arrayValueShape(args[0], name, optionBase) : undefined);
	state.held.push(heldOf(args[0]));
}

function remove(name: string, state: CollectionContents, arg: VbaToken[], base: Span, push: PushFn, indexOf: IndexOf): void {
	const span = { start: base.start + arg[0].start, end: base.start + arg[arg.length - 1].end };
	const index = indexOf(arg);
	if (index !== undefined) {
		if (state.items.length === 0 || index < 1 || index > state.items.length) {
			reportIndex(name, state, index, span, push);
			return;
		}
		state.items.splice(index - 1, 1);
		state.shapes.splice(index - 1, 1);
		state.held.splice(index - 1, 1);
		return;
	}
	const key = literalKey(arg);
	if (key === undefined) {
		// A variable index or key: one element fewer, which one unknown.
		state.items.pop();
		state.keysKnown = false;
		state.shapes = state.items.map(() => undefined);
		state.held = state.items.map(() => undefined);
		return;
	}
	if (reportKey(name, state, stringLiteralValue(arg[0].rawText), span, push)) {
		return;
	}
	const position = state.items.indexOf(key);
	if (position >= 0) {
		state.items.splice(position, 1);
		state.shapes.splice(position, 1);
		state.held.splice(position, 1);
	} else {
		state.items.pop();
		state.keysKnown = false;
		state.shapes = state.items.map(() => undefined);
		state.held = state.items.map(() => undefined);
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
