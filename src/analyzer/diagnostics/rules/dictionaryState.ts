// Rule family: a Scripting.Dictionary whose keys the code makes plain (issue
// #243). Measured in Excel 16.0 (build 20326, 2026-09-30); each compiles and
// raises every time it runs.
//
//  - `d.Add "a", 1` then `d.Add "a", 2` -> 457, the key is in use.
//  - `d.Remove "a"` with no such key -> 32811.
//  - `d.Keys()(3)` or `d.Items()(3)` past the last key -> 9. They are based
//    at 0.
//
// A Dictionary differs from a Collection, whose state collectionState.ts
// follows: its keys compare as written, "A" and "a" being two keys, and a
// number and a string being two; and reading `d("missing")` runs, adding the
// key with an Empty item. The local is followed from
// `Set d = CreateObject("Scripting.Dictionary")` or `New Scripting.Dictionary`.
// Anything else that names it ends what is known.
//
// Issue #349, measured in Excel 16.0 on 2026-10-02: `d.CompareMode = 1`
// after a key is in raises 5, and before it makes "k" and "K" one key;
// `d.Key("a") = "b"` raises 32811 with no "a" and 457 with "b" in use; an
// Array as the key of Add raises 5; and `d("k").Count` with no "k" adds the
// key with an Empty item, whose member raises 424.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import { statementLabelDeclaration } from '../../flow/procedureLabels';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { BodyNode, ModuleNode, Span } from '../../parser/nodes';
import { isLeafStatement } from '../../parser/nodes';
import type { PushFn } from '../analysisContext';
import { walkEnteringBlocks } from '../dataflow';
import { stringLiteralValue } from '../typeInference';
import {
	activeModuleMembers,
	matchParenFrom,
	setAssignmentTarget,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
} from '../walker';
import { namesIn } from './shared';

/** The keys of one Dictionary in the order they were added, each a typed literal. */
interface DictionaryKeys {
	keys: string[];
	/** Set by `CompareMode = 1`: string keys compare without case (issue #349). */
	textCompare?: boolean;
}

export function checkDictionaryState(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const states = new Map<string, DictionaryKeys>();
		const forget = (names: Iterable<string>): void => {
			for (const lower of names) {
				states.delete(lower);
			}
		};
		const visit = (node: BodyNode): void => {
			if (!isLeafStatement(node)) {
				return;
			}
			if (statementLabelDeclaration(source, node.span)) {
				states.clear();
			}
			if (node.kind === 'Statement' && node.singleLineIfBranches) {
				forget(namesIn(source, node.span));
				return;
			}
			const toks = statementTokensAfterLeadingLabel(source, node.span).filter((tok) => tok.kind !== 'comment');
			if (tokenText(toks[0]) === 'gosub') {
				states.clear();
				return;
			}
			const set = setAssignmentTarget(source, node.span);
			if (set) {
				const lower = set.name.toLowerCase();
				const value = toks.slice(toks.findIndex((tok) => tok.rawText === '=') + 1);
				forget(namesIn(source, node.span));
				if (createsDictionary(value)) {
					states.set(lower, { keys: [] });
				}
				return;
			}
			checkStatement(node.span, toks, states, push);
		};
		walkEnteringBlocks(source, member.body, (node) => activity?.isInactive(node.span) === true, visit, {
			snapshot: () => new Map([...states].map(([lower, state]) => [lower, { ...state, keys: [...state.keys] }])),
			restore: (saved) => {
				states.clear();
				for (const [lower, state] of saved) {
					states.set(lower, { ...state, keys: [...state.keys] });
				}
			},
			forget,
			touches: (stmt) => namesIn(source, stmt.span),
		});
	}
}

/** `CreateObject("Scripting.Dictionary")`, `VBA.CreateObject(...)` or `New Scripting.Dictionary`. */
function createsDictionary(value: readonly VbaToken[]): boolean {
	const text = value.map((tok) => tok.rawText).join('').toLowerCase();
	return /^(?:vba\.)?createobject\("scripting\.dictionary"\)$/.test(text) || text === 'newscripting.dictionary';
}

/** The literal key an argument names, typed so that "1" and 1 differ, or undefined. */
function literalKey(arg: readonly VbaToken[]): string | undefined {
	if (arg.length !== 1) {
		return undefined;
	}
	if (arg[0].kind === 'stringLiteral') {
		return `s:${stringLiteralValue(arg[0].rawText)}`;
	}
	return arg[0].kind === 'integerLiteral' && /^\d+$/.test(arg[0].rawText) ? `n:${Number(arg[0].rawText)}` : undefined;
}

function shownKey(key: string): string {
	return key.startsWith('s:') ? JSON.stringify(key.slice(2)) : key.slice(2);
}

/** The comma-separated arguments from `toks[from]` to the end, or within one pair of parentheses. */
function argumentsOf(toks: readonly VbaToken[], from: number, to = toks.length): VbaToken[][] {
	const out: VbaToken[][] = [];
	let current: VbaToken[] = [];
	let depth = 0;
	for (let i = from; i < to; i++) {
		const raw = toks[i].rawText;
		if (raw === '(') {
			depth++;
		} else if (raw === ')') {
			depth--;
		}
		if (raw === ',' && depth === 0) {
			out.push(current);
			current = [];
			continue;
		}
		current.push(toks[i]);
	}
	out.push(current);
	return out;
}

function checkStatement(base: Span, toks: readonly VbaToken[], states: Map<string, DictionaryKeys>, push: PushFn): void {
	const at = (first: VbaToken, last: VbaToken): Span => ({ start: base.start + first.start, end: base.start + last.end });
	const head = tokenName(toks[0])?.toLowerCase();
	const state = head ? states.get(head) : undefined;
	// Under `CompareMode = 1` "k" and "K" are one key.
	const keyOf = (arg: readonly VbaToken[], held: DictionaryKeys): string | undefined => {
		const key = literalKey(arg);
		return key && held.textCompare && key.startsWith('s:') ? `s:${key.slice(2).toLowerCase()}` : key;
	};
	const eq = toks.findIndex((tok) => tok.rawText === '=');
	// `d.CompareMode = 1`: refused once a key is in (issue #349, measured).
	if (state && toks[1]?.rawText === '.' && tokenText(toks[2]) === 'comparemode' && eq === 3) {
		const mode = tokenText(toks[4]);
		const text = toks.length === 5 && (mode === '1' || mode === 'vbtextcompare');
		const binary = toks.length === 5 && (mode === '0' || mode === 'vbbinarycompare');
		if (state.keys.length > 0) {
			push('collectionAddArgument', `Dictionary '${toks[0].rawText}' already holds ${state.keys.length} key${state.keys.length === 1 ? '' : 's'}, so its CompareMode cannot change. This will raise Run-time error '5': Invalid procedure call or argument.`, at(toks[2], toks[toks.length - 1]));
			return;
		}
		if (text || binary) {
			state.textCompare = text;
		} else {
			states.delete(head!);
		}
		return;
	}
	// `d.Key("a") = "b"` renames a key: 32811 with no "a", 457 with "b" in use.
	if (state && toks[1]?.rawText === '.' && tokenText(toks[2]) === 'key' && toks[3]?.rawText === '(' && eq === matchParenFrom(toks, 3) + 1) {
		const from = keyOf(toks.slice(4, eq - 1), state);
		const to = keyOf(toks.slice(eq + 1), state);
		if (!from || !to) {
			states.delete(head!);
			return;
		}
		if (!state.keys.includes(from)) {
			push('collectionKeyNotFound', `Dictionary '${toks[0].rawText}' holds no key ${shownKey(from)} to rename. This will raise Run-time error '32811': Application-defined or object-defined error.`, at(toks[4], toks[eq - 2]));
			return;
		}
		if (state.keys.includes(to) && to !== from) {
			push('collectionKeyInUse', `Dictionary '${toks[0].rawText}' already holds the key ${shownKey(to)}. This will raise Run-time error '457': This key is already associated with an element of this collection.`, at(toks[eq + 1], toks[toks.length - 1]));
			return;
		}
		state.keys[state.keys.indexOf(from)] = to;
		return;
	}
	// A method called as a statement: `d.Add k, v`, `d.Remove k`, `d.RemoveAll`.
	if (state && toks[1]?.rawText === '.' && !toks.some((tok) => tok.rawText === '=')) {
		const method = tokenText(toks[2]);
		const paren = toks[3]?.rawText === '(' && matchParenFrom(toks, 3) === toks.length - 1;
		const args = toks.length > 3 ? (paren ? argumentsOf(toks, 4, toks.length - 1) : argumentsOf(toks, 3)) : [];
		const key = args[0] ? keyOf(args[0], state) : undefined;
		// `d.Add Array(1), 1`: an array is no key (issue #349, measured).
		if (method === 'add' && args.length === 2 && tokenText(args[0][0]) === 'array' && args[0][1]?.rawText === '(' && matchParenFrom(args[0], 1) === args[0].length - 1) {
			push('collectionAddArgument', `The key of '${toks[0].rawText}.Add' is an array, which a Dictionary takes as no key. This will raise Run-time error '5': Invalid procedure call or argument.`, at(args[0][0], args[0][args[0].length - 1]));
			return;
		}
		if (method === 'add' && args.length === 2 && key) {
			if (state.keys.includes(key)) {
				push('collectionKeyInUse', `Dictionary '${toks[0].rawText}' already holds the key ${shownKey(key)}. This will raise Run-time error '457': This key is already associated with an element of this collection.`, at(args[0][0], args[0][args[0].length - 1]));
				return;
			}
			state.keys.push(key);
			return;
		}
		if (method === 'remove' && args.length === 1 && key) {
			const index = state.keys.indexOf(key);
			if (index < 0) {
				push('collectionKeyNotFound', `Dictionary '${toks[0].rawText}' holds no key ${shownKey(key)} to remove. This will raise Run-time error '32811': Application-defined or object-defined error.`, at(args[0][0], args[0][args[0].length - 1]));
				return;
			}
			state.keys.splice(index, 1);
			return;
		}
		if (method === 'removeall' && args.length === 0) {
			state.keys = [];
			return;
		}
	}
	// Reads, anywhere in the statement: `d.Keys()(n)`, `d.Items()(n)`, `d(k)`, `d.Item(k)`.
	for (let i = 0; i < toks.length; i++) {
		const lower = tokenName(toks[i])?.toLowerCase();
		const read = lower ? states.get(lower) : undefined;
		if (!read || toks[i - 1]?.rawText === '.' || toks[i - 1]?.rawText === '!') {
			continue;
		}
		const member = toks[i + 1]?.rawText === '.' ? tokenText(toks[i + 2]) : '';
		if ((member === 'keys' || member === 'items') && toks[i + 3]?.rawText === '(' && toks[i + 4]?.rawText === ')' && toks[i + 5]?.rawText === '(') {
			const close = matchParenFrom(toks, i + 5);
			const index = close === i + 7 && toks[i + 6].kind === 'integerLiteral' && /^\d+$/.test(toks[i + 6].rawText) ? Number(toks[i + 6].rawText) : undefined;
			if (index !== undefined && index >= read.keys.length) {
				const held = read.keys.length === 0 ? 'holds no keys' : `holds ${read.keys.length} key${read.keys.length === 1 ? '' : 's'}, indexed 0 to ${read.keys.length - 1}`;
				push('collectionIndexOutOfRange', `Dictionary '${toks[i].rawText}' ${held} here; ${index} is outside that. This will raise Run-time error '9': Subscript out of range.`, at(toks[i + 6], toks[i + 6]));
			}
			continue;
		}
		if (member === 'count' || member === 'exists') {
			continue;
		}
		// `d(k)` and `d.Item(k)` add a key they do not find, read or written.
		const open = toks[i + 1]?.rawText === '(' ? i + 1 : member === 'item' && toks[i + 3]?.rawText === '(' ? i + 3 : -1;
		const close = open >= 0 ? matchParenFrom(toks, open) : -1;
		const key = close > open + 1 ? keyOf(toks.slice(open + 1, close), read) : undefined;
		if (key) {
			if (!read.keys.includes(key)) {
				// `d("k").Count` with no "k": the read adds it with an Empty
				// item, which has no members (issue #349, measured).
				if (toks[close + 1]?.rawText === '.') {
					push('variantValueMisuse', `Dictionary '${toks[i].rawText}' holds no key ${shownKey(key)}, so the read adds it with an Empty item, which has no ${toks[close + 2]?.rawText ?? 'member'}. This will raise Run-time error '424': Object required.`, at(toks[i], toks[close]));
				}
				read.keys.push(key);
			}
			continue;
		}
		// Anything else may change it.
		states.delete(lower!);
	}
}
