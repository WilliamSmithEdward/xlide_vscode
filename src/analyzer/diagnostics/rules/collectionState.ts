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
import { jumpTargetLabelDeclaration } from '../../flow/procedureLabels';
import { procedureSymbolFor, type PushFn } from '../analysisContext';
import { counterText, loopCountersAt, numericCounterPasses, type LoopCounter } from '../loopCounters';
import { calleeMemberCalls, type CalleeMemberCalls } from '../calleeArguments';
import { isKnownScalarType, knownLocalLiteralValuesAt, normalizeType, procedureIntegerConstantLookup, stringLiteralValue, unreachableStatementsIn, withKnownLocals } from '../typeInference';
import {
	activeModuleMembers,
	blockHeaderLineSpan,
	forEachVariableGroup,
	matchParenFrom,
	rawExpressionTokens,
	setAssignmentTarget,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
} from '../walker';
import { bankersRound, nameMentions, namesIn } from './shared';
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

/** What `c(1)`, `c("k")` or `c.Item(1)`, the whole of a value, reads from a tracked collection. */
function heldItemRead(value: readonly VbaToken[], states: ReadonlyMap<string, CollectionContents>): { held: Held } | undefined {
	const toks = value.filter((tok) => tok.kind !== 'comment');
	const contents = states.get(tokenName(toks[0])?.toLowerCase() ?? '');
	const open = toks[1]?.rawText === '(' ? 1 : toks[1]?.rawText === '.' && tokenText(toks[2]) === 'item' && toks[3]?.rawText === '(' ? 3 : -1;
	if (!contents || contents.stale || open < 0 || matchParenFrom(toks, open) !== toks.length - 1) {
		return undefined;
	}
	const arg = toks.slice(open + 1, toks.length - 1);
	const key = literalKeyText(arg);
	const index = literalIndex(arg) ?? (key !== undefined && contents.keysKnown && contents.items.includes(key) ? contents.items.indexOf(key) + 1 : undefined);
	return index !== undefined && index >= 1 && index <= contents.held.length ? { held: contents.held[index - 1] } : undefined;
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
	const calleeCalls = calleeMemberCalls(source);
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
		if (states.size === 0 && autoInstanced.plainLocals.size === 0 && !/\bWith[ \t]+New[ \t]+Collection\b/i.test(source.slice(member.span.start, member.span.end))) {
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
				// `If x Then .Add 20` inside `With c` may change c (issue #584).
				const within = withSubjects[withSubjects.length - 1];
				if (within && reachesSubject(statementTokensAfterLeadingLabel(source, node.span), within)) {
					forgetCollection(states, subjectName(within));
				}
				return;
			}
			const own = statementTokensAfterLeadingLabel(source, node.span);
			if (own.length === 0) {
				return;
			}
			// Inside `With c` or `With New Collection`, `.Item(2)` is the
			// subject's (issue #295, measured in Excel 16.0).
			const subject = withSubjects[withSubjects.length - 1];
			const toks = subject ? withReceiver(own, subject) : own;
			// A label may be reached from anywhere; a GoSub may run any statement.
			if (jumpTargetLabelDeclaration(source, node.span) || tokenText(toks[0]) === 'gosub') {
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
				const item = heldItemRead(value, states);
				if (item && (item.held === 'number' || item.held === 'string')) {
					const first = value[0];
					const last = value[value.length - 1];
					push('variantValueMisuse', `'${value.map((tok) => tok.rawText).join('')}' holds a ${item.held}, not an object, so Set has nothing to assign. This will raise Run-time error '424': Object required.`, { start: node.span.start + first.start, end: node.span.start + last.end });
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
			// `c(1.6)` rounds to 2, half to even (issue #349, measured in Excel 16.0).
			const indexOf = (arg: readonly VbaToken[]): number | undefined => literalIndex(arg)
				?? (arg.length === 1 && arg[0].kind === 'floatLiteral' && Number.isFinite(Number(arg[0].rawText.replace(/[!#@]$/, ''))) ? bankersRound(Number(arg[0].rawText.replace(/[!#@]$/, ''))) : undefined)
				?? (lookup ? evaluateIntegerConstantExpression(arg.map((tok) => tok.rawText).join(' '), lookup) : undefined);
			const held = valuesAt?.(node);
			const keyOf: KeyOf = (arg) => {
				const local = arg.length === 1 ? tokenName(arg[0])?.toLowerCase() : undefined;
				const value = local ? held?.get(local) : undefined;
				return literalKeyText(arg) ?? (value?.kind === 'string' && !value.contentMutated ? value.value as string : undefined);
			};
			// `R1 c` where R1 only adds to or removes from its parameter: c
			// changes as those calls change it (issue #685).
			const replays = replayedCalls(toks, states, calleeCalls);
			if (replays) {
				for (const [lower, calls] of replays) {
					for (const call of calls) {
						let raised = false;
						checkStatement(node.span, call, states, () => { raised = true; }, isEmpty, indexOf, optionBase, lookup, scalarLocal, keyOf);
						if (raised || !states.has(lower)) {
							forgetCollection(states, lower);
							break;
						}
					}
				}
				return;
			}
			checkStatement(node.span, toks, states, push, isEmpty, indexOf, optionBase, lookup, scalarLocal, keyOf);
		};
		const leaves = new Map<BodyNode, { after: Map<string, CollectionContents>; aliases: string[] }>();
		// The subject of each With block the walk is in: a tracked local's
		// name, its element `c(1)`, `New Collection` for a new one, or
		// undefined for anything else.
		const withSubjects: Array<WithSubject | undefined> = [];
		const enterWith = (node: BodyNode): void => {
			const header = statementTokensAfterLeadingLabel(source, blockHeaderLineSpan(source, node.span)).filter((tok) => tok.kind !== 'comment');
			const name = header.length === 2 ? tokenName(header[1]) : undefined;
			const element = header.length >= 5 && header[2].rawText === '(' && matchParenFrom([...header], 2) === header.length - 1 ? tokenName(header[1]) : undefined;
			if (name && states.has(name.toLowerCase())) {
				withSubjects.push(name);
			} else if (element && states.has(element.toLowerCase())) {
				// `With c(1)` on a collection whose element is a collection:
				// `.Item(2)` reads as `c(1).Item(2)` (issue #295).
				withSubjects.push(header.slice(1));
			} else if (header.length === 3 && tokenText(header[1]) === 'new' && tokenText(header[2]) === 'collection') {
				states.set(NEW_WITH_SUBJECT.toLowerCase(), emptyContents());
				withSubjects.push(NEW_WITH_SUBJECT);
			} else {
				withSubjects.push(undefined);
			}
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
			// `With c` reads c and changes nothing; its body's lines say what they do,
			// a `.Add` among them naming the subject (issue #584).
			touches: (stmt) => {
				const toks = statementTokensAfterLeadingLabel(source, stmt.span).filter((tok) => tok.kind !== 'comment');
				// `With c(1)` reads one element, by a literal, and changes nothing.
				if (tokenText(toks[0]) === 'with' && (toks.length === 2
					|| (toks.length === 5 && toks[2].rawText === '(' && (toks[3].kind === 'integerLiteral' || toks[3].kind === 'stringLiteral') && toks[4].rawText === ')'))) {
					return new Set<string>();
				}
				const names = namesIn(source, stmt.span);
				const within = withSubjects[withSubjects.length - 1];
				return within && reachesSubject(toks, within) ? new Set([...names, subjectName(within)]) : names;
			},
			withBodyRunsThrough: true,
			// A counted loop that removes or reads by its counter (issue #263),
			// or fills or empties a collection (issue #350).
			enter: (node) => {
				if (node.kind === 'WithBlock') {
					enterWith(node);
				}
				checkScalarElements(source, node, states, symbols ? procedureSymbolFor(symbols, member)?.children ?? [] : [], push);
				simulateCountedLoop(source, node, states, push, activity);
				const after = unreachable?.has(node) ? undefined : simulateFillingLoop(source, node, states, push, activity);
				if (after) {
					// Another name for the same collection, `Set o = c` above,
					// is not followed past the loop.
					const aliases = [...states].filter(([lower, contents]) => !after.has(lower) && [...after.keys()].some((name) => states.get(name) === contents)).map(([lower]) => lower);
					leaves.set(node, { after, aliases });
				}
			},
			exit: (node) => {
				if (node.kind === 'WithBlock' && withSubjects.pop() === NEW_WITH_SUBJECT) {
					forgetCollection(states, NEW_WITH_SUBJECT.toLowerCase());
				}
				const left = leaves.get(node);
				for (const [lower, contents] of left?.after ?? []) {
					states.set(lower, contents);
				}
				for (const lower of left?.aliases ?? []) {
					forgetCollection(states, lower);
				}
			},
		});
	}
}

/**
 * The member calls a call statement's callee makes on the tracked objects
 * passed to it, when each tracked name the statement mentions is one of
 * them, passed whole. Undefined otherwise, and the statement is judged as
 * any other (issue #685).
 */
export function replayedCalls(
	toks: readonly VbaToken[],
	states: ReadonlyMap<string, unknown>,
	calleeCalls: CalleeMemberCalls,
): ReadonlyMap<string, readonly (readonly VbaToken[])[]> | undefined {
	const mentioned = toks.filter((tok, i) => states.has(tokenName(tok)?.toLowerCase() ?? '') && toks[i - 1]?.rawText !== '.');
	if (mentioned.length === 0) {
		return undefined;
	}
	const replays = calleeCalls(toks);
	return mentioned.every((tok) => replays.has(tokenName(tok)!.toLowerCase())) ? replays : undefined;
}

/** A With block's subject: a name, or the tokens of an element, `c ( 1 )`. */
type WithSubject = string | readonly VbaToken[];

/** The lowercased local a With subject is, or is an element of. */
function subjectName(subject: WithSubject): string {
	return (typeof subject === 'string' ? subject : subject[0].rawText).toLowerCase();
}

/** Whether a statement inside `With subject` reaches the subject by a leading dot. */
function reachesSubject(toks: readonly VbaToken[], subject: WithSubject): boolean {
	return withReceiver(toks.filter((tok) => tok.kind !== 'comment'), subject).length > toks.filter((tok) => tok.kind !== 'comment').length;
}

/**
 * `For Each v In c` with every element of c a number or a string: v holds
 * one, and the body's first line to name v, `v Is Nothing` or `v.Count`,
 * raises 424 (issue #612, measured in Excel 16.0).
 */
function checkScalarElements(source: string, node: BodyNode, states: ReadonlyMap<string, CollectionContents>, locals: readonly VbaSymbol[], push: PushFn): void {
	if (node.kind !== 'ForBlock' || !node.each || !node.controlVariable || !node.sourceExpression) {
		return;
	}
	const lower = node.controlVariable.toLowerCase();
	const contents = states.get(node.sourceExpression.trim().toLowerCase());
	const local = locals.find((child) => child.name.toLowerCase() === lower);
	const type = normalizeType(local?.asType);
	if (!contents || contents.stale || contents.held.length === 0 || !contents.held.every((held) => held === 'number' || held === 'string')
		|| local?.kind !== 'localVariable' || (type !== undefined && type !== 'variant')) {
		return;
	}
	for (const child of node.body) {
		if (!isLeafStatement(child)) {
			if (new RegExp(`\\b${lower}\\b`, 'i').test(source.slice(child.span.start, child.span.end))) {
				return;
			}
			continue;
		}
		const toks = statementTokensAfterLeadingLabel(source, child.span).filter((tok) => tok.kind !== 'comment');
		const at = toks.findIndex((tok, i) => tokenName(tok)?.toLowerCase() === lower && toks[i - 1]?.rawText !== '.');
		if (at < 0) {
			continue;
		}
		const isOperand = (tokenText(toks[at + 1]) === 'is' && tokenText(toks[at - 1]) !== 'typeof') || (tokenText(toks[at - 1]) === 'is' && tokenText(toks[at - 2]) !== 'typeof');
		const member = toks[at + 1]?.rawText === '.' && tokenName(toks[at + 2]) !== undefined;
		if (isOperand || member) {
			push('variantValueMisuse', `'${toks[at].rawText}' holds an element of '${node.sourceExpression.trim()}', each a number or a string, not an object${member ? ' with members' : ' for Is to compare'}. This will raise Run-time error '424': Object required.`, { start: child.span.start + toks[at].start, end: child.span.start + toks[at].end });
		}
		return;
	}
}

/** The name a `With New Collection` block's collection is followed under. */
const NEW_WITH_SUBJECT = 'New Collection';

/**
 * A statement's tokens with the With subject before each member the block
 * reaches by a leading dot: `.Item(2)` reads as `c.Item(2)`.
 */
export function withReceiver(toks: readonly VbaToken[], subject: WithSubject): VbaToken[] {
	const out: VbaToken[] = [];
	toks.forEach((tok, i) => {
		const before = toks[i - 1];
		const leading = tok.rawText === '.' && (!before || (before.kind !== 'identifier' && before.kind !== 'bracketedIdentifier' && before.rawText !== ')' && !(before.kind === 'keyword' && tokenText(before) === 'me')));
		if (leading && tokenName(toks[i + 1])) {
			if (typeof subject === 'string') {
				out.push({ ...tok, kind: 'identifier', rawText: subject, end: tok.start });
			} else {
				out.push(...subject.map((part) => ({ ...part, start: tok.start, end: tok.start })));
			}
		}
		out.push(tok);
	});
	return out;
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
		if (['exit', 'goto', 'gosub', 'resume', 'return', 'end', 'on', 'stop'].includes(head) || jumpTargetLabelDeclaration(source, stmt.span)) {
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

/**
 * `For i = 1 To 3: c.Add i: Next` (issue #350, measured in Excel 16.0): a
 * For loop with literal bounds whose body is plain statements, each either
 * `c.Add item[, key]` or `c.Remove n` on a tracked collection or one that
 * names none, is run pass by pass. A literal key added on a second pass
 * raises 457 there. What the loop leaves, an empty collection after a loop
 * of no pass included, is returned for after it; a key the code builds
 * leaves the keys unknown. Undefined when the loop cannot be followed.
 */
function simulateFillingLoop(
	source: string,
	node: BodyNode,
	states: ReadonlyMap<string, CollectionContents>,
	push: PushFn,
	activity: ConditionalActivityTracker | undefined,
): Map<string, CollectionContents> | undefined {
	if (node.kind !== 'ForBlock' || !node.controlVariable || states.size === 0) {
		return undefined;
	}
	const counter = node.controlVariable.toLowerCase();
	// The values the loop variable takes, pass by pass: a counted For's, or
	// the elements of a literal Split or Array a For Each steps through.
	const values: Array<number | string> = [];
	if (node.each) {
		const elements = literalElements(node.sourceExpression ?? '');
		if (!elements) {
			return undefined;
		}
		values.push(...elements);
	} else {
		const header = statementTokensAfterLeadingLabel(source, blockHeaderLineSpan(source, node.span)).filter((tok) => tok.kind !== 'comment');
		const eq = header.findIndex((tok) => tok.rawText === '=');
		const to = header.findIndex((tok) => tokenText(tok) === 'to');
		const stepAt = header.findIndex((tok) => tokenText(tok) === 'step');
		const start = eq > 0 && to > eq ? literalIndex(header.slice(eq + 1, to)) : undefined;
		const limit = to > 0 ? literalIndex(header.slice(to + 1, stepAt > 0 ? stepAt : header.length)) : undefined;
		const step = stepAt > 0 ? literalIndex(header.slice(stepAt + 1)) : 1;
		if (start === undefined || limit === undefined || !step) {
			return undefined;
		}
		for (let value = start; step > 0 ? value <= limit : value >= limit; value += step) {
			if (values.length >= MAX_SIMULATED_PASSES) {
				return undefined;
			}
			values.push(value);
		}
	}
	interface Change { name: string; display: string; add?: { key?: string; keyToks?: readonly VbaToken[]; keyBuilt: boolean; keyIsVariable?: boolean }; remove?: number; base: number }
	const changes: Change[] = [];
	for (const stmt of node.body) {
		if (activity?.isInactive(stmt.span)) {
			continue;
		}
		if (!isLeafStatement(stmt) || (stmt.kind === 'Statement' && stmt.singleLineIfBranches) || jumpTargetLabelDeclaration(source, stmt.span)) {
			return undefined;
		}
		const toks = statementTokensAfterLeadingLabel(source, stmt.span).filter((tok) => tok.kind !== 'comment');
		if (['exit', 'goto', 'gosub', 'resume', 'return', 'end', 'on', 'stop'].includes(tokenText(toks[0]))) {
			return undefined;
		}
		const name = tokenName(toks[0])?.toLowerCase();
		const mentions = toks.some((tok, k) => states.has(tokenName(tok)?.toLowerCase() ?? '') && toks[k - 1]?.rawText !== '.');
		if (!name || !states.has(name) || toks[1]?.rawText !== '.') {
			if (mentions || toks.some((tok) => tokenName(tok)?.toLowerCase() === counter && tok === toks[0])) {
				return undefined;
			}
			continue;
		}
		const member = tokenText(toks[2]);
		const args = argumentsAfter(toks, 3);
		if (member === 'add' && args.length >= 1 && args.length <= 2 && !args.some((arg) => arg[1]?.rawText === ':=')) {
			const keyToks = args[1];
			const key = keyToks ? literalKey(keyToks) : undefined;
			// `c.Add p, p` in a For Each: the key is the element of the pass.
			const keyIsVariable = node.each && keyToks?.length === 1 && tokenName(keyToks[0])?.toLowerCase() === counter;
			changes.push({ name, display: toks[0].rawText, base: stmt.span.start, add: { key, keyToks, keyBuilt: !keyIsVariable && keyToks !== undefined && keyToks.length > 0 && key === undefined, keyIsVariable } });
			continue;
		}
		const index = member === 'remove' && args.length === 1 ? literalIndex(args[0]) : undefined;
		if (index === undefined) {
			return undefined;
		}
		changes.push({ name, display: toks[0].rawText, base: stmt.span.start, remove: index });
	}
	if (changes.length === 0) {
		return undefined;
	}
	const after = new Map<string, CollectionContents>();
	for (const change of changes) {
		if (!after.has(change.name)) {
			const contents = states.get(change.name)!;
			after.set(change.name, { items: [...contents.items], keysKnown: contents.keysKnown, shapes: [...contents.shapes], held: [...contents.held] });
		}
	}
	let passes = 0;
	for (const value of values) {
		passes++;
		for (const change of changes) {
			const contents = after.get(change.name)!;
			if (change.remove !== undefined) {
				if (change.remove < 1 || change.remove > contents.items.length) {
					return undefined; // the walk into the block reports a first pass
				}
				contents.items.splice(change.remove - 1, 1);
				contents.shapes.splice(change.remove - 1, 1);
				contents.held.splice(change.remove - 1, 1);
				continue;
			}
			const key = change.add!.keyIsVariable ? String(value).toLowerCase() : change.add!.key;
			if (key !== undefined && contents.keysKnown && contents.items.includes(key)) {
				if (passes > 1) {
					const keyToks = change.add!.keyToks!;
					push('collectionKeyInUse', `On the pass of the For loop where '${node.controlVariable}' is ${typeof value === 'string' ? JSON.stringify(value) : value}, '${change.display}' already has an element with the key ${keyToks[0].rawText} from an earlier pass. This will raise Run-time error '457': This key is already associated with an element of this collection.`, { start: change.base + keyToks[0].start, end: change.base + keyToks[keyToks.length - 1].end });
				}
				return undefined;
			}
			contents.items.push(key);
			contents.shapes.push(undefined);
			contents.held.push(undefined);
			if (change.add!.keyBuilt) {
				contents.keysKnown = false;
			}
		}
	}
	return after;
}

/** The elements of `Split("a,b", ",")` or `Array("a", "b")` written with literals, as Strings (issue #350). */
function literalElements(expression: string): string[] | undefined {
	const toks = rawExpressionTokens(expression).filter((tok) => tok.kind !== 'comment');
	const callee = tokenText(toks[0]);
	if ((callee !== 'split' && callee !== 'array') || toks[1]?.rawText !== '(' || matchParenFrom(toks, 1) !== toks.length - 1) {
		return undefined;
	}
	const args = argumentsAfter(toks, 1);
	if (!args.every((arg) => arg.length === 1 && arg[0].kind === 'stringLiteral')) {
		return undefined;
	}
	const texts = args.map((arg) => stringLiteralValue(arg[0].rawText));
	if (callee === 'array') {
		return texts;
	}
	if (texts.length < 1 || texts.length > 2 || texts[1] === '' || texts[0] === '') {
		return undefined;
	}
	return texts[0].split(texts[1] ?? ' ');
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
			// An Object is followed from `Set o = New Collection` on, as a
			// Collection is: `o(1)` on it empty raises 5 (issue #415).
			if (decl.isArray || (type !== 'collection' && !(type === 'object' && !decl.isNew))) {
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
	keyOf: KeyOf = literalKeyText,
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
		// `c.Add New Collection`: an empty Collection no name holds (issue #306).
		if (item.length === 2 && tokenText(item[0]) === 'new' && tokenText(item[1]) === 'collection') {
			return emptyContents();
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
			if (close > i + 2 && checkRead(lower, state, toks.slice(i + 2, close), at(i + 2, close - 1), push, indexOf, keyOf)) {
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
			if (itemClose > i + 4 && checkRead(lower, state, toks.slice(i + 4, itemClose), at(i + 4, itemClose - 1), push, indexOf, keyOf)) {
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
	// `Set c(1) = o`: Item has no Property Set, so the Set reaches what the
	// item holds: 424 on a value, 438 on an object (issue #306, measured in
	// Excel 16.0).
	if (tokenText(toks[0]) === 'set' && nameAt === 1 && after === '=') {
		if (held === 'number' || held === 'string') {
			ctx.push('variantValueMisuse', `'${shown}' holds a ${held}, and a Collection's item cannot be replaced in place: Item has no Property Set, so the Set reaches the ${held}. This will raise Run-time error '424': Object required.`, spanOf(nameAt, close));
		} else {
			ctx.push('runtimeMemberNotFound', `'${shown}' is read through Item, which a Set cannot write: a Collection's item cannot be replaced in place. This will raise Run-time error '438': Object doesn't support this property or method.`, spanOf(nameAt, close));
		}
		return;
	}
	if (held === 'number' || held === 'string') {
		// `c(1) = 5`: a Let into the item, which only an object's default
		// member could take (issue #305, measured in Excel 16.0).
		if (after === '=' && nameAt === ctx.first) {
			ctx.push('variantValueMisuse', `'${shown}' holds a ${held}, and a Collection's item cannot be replaced in place: only an object item takes a value through its default member. This will raise Run-time error '424': Object required.`, spanOf(nameAt, close));
		} else if (after === '(') {
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
function checkRead(name: string, state: CollectionContents, arg: readonly VbaToken[], span: Span, push: PushFn, indexOf: IndexOf, keyOf: KeyOf = literalKeyText): boolean {
	const index = indexOf(arg);
	if (index !== undefined) {
		reportIndex(name, state, index, span, push);
		return true;
	}
	const key = keyOf(arg);
	if (key !== undefined) {
		reportKey(name, state, key, span, push);
		return true;
	}
	return false;
}

/** The key an argument names: a string literal, or a String local known to hold one (issue #346). */
type KeyOf = (arg: readonly VbaToken[]) => string | undefined;

function literalKeyText(arg: readonly VbaToken[]): string | undefined {
	return arg.length === 1 && arg[0].kind === 'stringLiteral' ? stringLiteralValue(arg[0].rawText) : undefined;
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
	// `k = 5` then `c.Add 1, k`: a local known to hold a number (issue #349).
	const heldNumber = keyLiteral.length === 1 && tokenName(keyLiteral[0]) ? indexOf(keyLiteral) : undefined;
	if (heldNumber !== undefined) {
		return { rule: 'collectionAddArgument', message: `The key of '${name}.Add' is '${keyLiteral[0].rawText}', which holds the number ${heldNumber}, not a string. This will raise Run-time error '13': Type mismatch.`, span: spanOf(keyLiteral) };
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
	// `c.Add 2, "b", "zz"`: Before names a key no element has (issue #349).
	const positionKey = literalKey(position);
	if (positionKey !== undefined && state.keysKnown && !state.items.includes(positionKey)) {
		return { rule: 'collectionAddArgument', message: `No element of '${name}' was added with the key ${position[0].rawText}, which ${before ? 'Before' : 'After'} names. This will raise Run-time error '5': Invalid procedure call or argument.`, span: spanOf(position) };
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
