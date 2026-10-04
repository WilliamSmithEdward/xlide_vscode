// What object a local is known to hold at a statement, and the classes of
// the items a Collection local holds (issue #246). Read from `Set x = New C`,
// from `Set x = y` with y known, and from `c.Add New C`. Anything else that
// names a local whole, which may pass it ByRef or replace it, ends what is
// known of it; a member call or an index read (`x.Foo`, `c(1)`, `c.Count`)
// does not. Blocks are entered as issue #237 enters them.

import type { ConditionalActivityTracker } from '../conditional/conditionalCompilation';
import { jumpTargetLabelDeclaration } from '../flow/procedureLabels';
import { splitTopLevelTokenGroups } from '../lexer/tokenHelpers';
import type { VbaToken } from '../lexer/tokenKinds';
import type { BodyNode, ProcedureNode } from '../parser/nodes';
import { isLeafStatement } from '../parser/nodes';
import type { buildModuleSymbols } from '../symbols/buildModuleSymbols';
import { procedureSymbolFor } from './analysisContext';
import { walkEnteringBlocks } from './dataflow';
import { namesIn } from './rules/shared';
import { normalizeType } from './typeInference';
import { setAssignmentTarget, statementTokensAfterLeadingLabel, tokenName, tokenText } from './walker';

export interface HeldObjects {
	/** The class each local holds, by lowercased name, as written after New. */
	classes: ReadonlyMap<string, string>;
	/** The classes of a Collection local's items, in order. */
	items: ReadonlyMap<string, readonly string[]>;
}

interface State {
	classes: Map<string, string>;
	items: Map<string, string[]>;
}

const NOTHING_HELD: HeldObjects = { classes: new Map(), items: new Map() };

/** The class recorded for ActiveSheet, which may be a Worksheet or a Chart. */
export const ACTIVE_SHEET_HELD = 'Worksheet or Chart';

/** The item class recorded for a number or string a Collection holds: no object. */
export const HELD_VALUE = '(value)';

/**
 * `Set x = Application`, `Set x = ActiveWorkbook.Names` and, in Word,
 * `Set x = ActiveDocument`: the host's own object, which is never Nothing
 * (issues #415 and #438).
 */
function hostObjectHeld(value: readonly VbaToken[], declared: ReadonlySet<string>): string | undefined {
	const last = tokenText(value[value.length - 1]);
	if (value.length === 1 && last === 'application' && !declared.has('application')) {
		return 'Application';
	}
	// Word's own Document (issue #438).
	if (value.length === 1 && (last === 'activedocument' || last === 'thisdocument') && !declared.has(last)) {
		return 'Document';
	}
	if (last === 'names' && (value.length === 1 ? !declared.has('names') : value[value.length - 2]?.rawText === '.')) {
		return 'Names';
	}
	// Excel's ActiveSheet, whichever kind of sheet it is: a Collection
	// parameter refuses it with 13 (issue #685).
	if (value.length === 1 && last === 'activesheet' && !declared.has(last)) {
		return ACTIVE_SHEET_HELD;
	}
	return undefined;
}

/** The class each ProgID CreateObject makes, by lowercased ProgID (issue #685). */
const PROGID_CLASSES: Readonly<Record<string, string>> = {
	'scripting.dictionary': 'Scripting.Dictionary',
	'scripting.filesystemobject': 'Scripting.FileSystemObject',
};

/** `Set d = CreateObject("Scripting.Dictionary")`: a Dictionary, which a Collection parameter refuses with 13 (issue #685). */
function createdByProgId(value: readonly VbaToken[], declared: ReadonlySet<string>): string | undefined {
	const at = tokenText(value[0]) === 'vba' && value[1]?.rawText === '.' ? 2 : 0;
	if (tokenText(value[at]) !== 'createobject' || declared.has('createobject') || value[at + 1]?.rawText !== '(' || value[at + 2]?.kind !== 'stringLiteral'
		|| value[at + 3]?.rawText !== ')' || value.length !== at + 4) {
		return undefined;
	}
	return PROGID_CLASSES[value[at + 2].rawText.slice(1, -1).toLowerCase()];
}

/** Members that read a Collection without changing it. */
const COLLECTION_READS: ReadonlySet<string> = new Set(['count', 'item']);

type HeldFacts = (node: BodyNode) => HeldObjects;
interface CachedHeldFacts {
	source: string;
	activity: ConditionalActivityTracker | undefined;
	facts: HeldFacts;
}
// Symbol and procedure snapshots own the cache lifetime. Keep only the last
// source/activity pair per procedure, with no chain of older source versions.
const HELD_FACTS = new WeakMap<
	ReturnType<typeof buildModuleSymbols>,
	WeakMap<ProcedureNode, CachedHeldFacts>
>();

/**
 * What each statement, and each block as it is entered, sees. A statement
 * the walk never reached sees nothing.
 */
export function heldObjectsAt(
	source: string,
	proc: ProcedureNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
	/** The class of any other value a Set gives, `Set o = Range("A1").Font`, where the caller can tell (issue #685). */
	classOfValue?: (value: readonly VbaToken[], offset: number) => string | undefined,
): HeldFacts {
	// Callbacks can capture changing host/type contexts: preserve their original
	// evaluation behavior even when the function identity has not changed.
	if (classOfValue) { return collectHeldObjects(source, proc, symbols, activity, classOfValue); }
	let byProcedure = HELD_FACTS.get(symbols);
	if (!byProcedure) {
		byProcedure = new WeakMap();
		HELD_FACTS.set(symbols, byProcedure);
	}
	const cached = byProcedure.get(proc);
	if (cached?.source === source && cached.activity === activity) {
		cached.source = source;
		return cached.facts;
	}
	const facts = collectHeldObjects(source, proc, symbols, activity);
	byProcedure.set(proc, {source, activity, facts});
	return facts;
}

function collectHeldObjects(
	source: string,
	proc: ProcedureNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
	classOfValue?: (value: readonly VbaToken[], offset: number) => string | undefined,
): HeldFacts {
	const seen = new Map<BodyNode, HeldObjects>();
	const state: State = { classes: new Map(), items: new Map() };
	// The names a local or parameter takes, which hide a host's global.
	const declared = new Set([...(procedureSymbolFor(symbols, proc)?.children ?? []).map((child) => child.name.toLowerCase()), ...proc.params.map((param) => param.name.toLowerCase())]);
	// `Dim c As New Collection` holds an empty one from the start.
	for (const child of procedureSymbolFor(symbols, proc)?.children ?? []) {
		if (child.kind === 'localVariable' && child.isAutoInstantiated && !child.isArray && child.asType) {
			state.classes.set(child.name.toLowerCase(), child.asType);
			if (normalizeType(child.asType) === 'collection') {
				state.items.set(child.name.toLowerCase(), []);
			}
		}
	}
	// Maps and item arrays are copied; reuse the snapshot until tracked state changes.
	let currentSnapshot: State | undefined;
	const snapshot = (): State => currentSnapshot ??= { classes: new Map(state.classes), items: new Map([...state.items].map(([k, v]) => [k, [...v]])) };
	const forget = (names: Iterable<string>): void => {
		for (const lower of names) {
			const removedClass = state.classes.delete(lower);
			const removedItems = state.items.delete(lower);
			if (removedClass || removedItems) { currentSnapshot = undefined; }
		}
	};
	const visit = (node: BodyNode): void => {
		if (!isLeafStatement(node)) {
			return;
		}
		const toks = statementTokensAfterLeadingLabel(source, node.span);
		if (jumpTargetLabelDeclaration(source, node.span) || tokenText(toks[0]) === 'gosub') {
			forget([...state.classes.keys(), ...state.items.keys()]);
		}
		if (state.classes.size > 0 || state.items.size > 0) {
			seen.set(node, snapshot());
		}
		if (node.kind === 'Statement' && node.singleLineIfBranches) {
			forget(namesIn(source, node.span));
			return;
		}
		const set = setAssignmentTarget(source, node.span);
		if (set) {
			const lower = set.name.toLowerCase();
			const value = toks.slice(toks.findIndex((tok) => tok.rawText === '=') + 1);
			const from = value.length === 1 ? tokenName(value[0])?.toLowerCase() : undefined;
			const created = value.length === 2 && tokenText(value[0]) === 'new' ? tokenName(value[1]) : undefined;
			const held = created ?? (from ? state.classes.get(from) : undefined) ?? hostObjectHeld(value, declared) ?? createdByProgId(value, declared)
				?? classOfValue?.(value, node.span.start);
			// The value's own holder may now change it unseen.
			forget([lower, ...namesIn(source, node.span)].filter((name) => name !== from || !held));
			if (from && held) {
				if (state.items.delete(from)) { currentSnapshot = undefined; }
			}
			if (held) {
				currentSnapshot = undefined;
				state.classes.set(lower, held);
				if (created && normalizeType(created) === 'collection') {
					state.items.set(lower, []);
				}
			}
			return;
		}
		// `c.Add New Flat1` puts a Flat1 at the end; Before:=1 or After:=1, a
		// whole number, puts it there (issue #356). A key or an expression for
		// either leaves the order unknown.
		const head = tokenName(toks[0])?.toLowerCase();
		if (head && state.items.has(head) && toks[1]?.rawText === '.' && tokenText(toks[2]) === 'add'
			&& tokenText(toks[3]) === 'new' && tokenName(toks[4]) && (toks[5] === undefined || toks[5].rawText === ',')) {
			const items = state.items.get(head)!;
			const at = addPosition(splitTopLevelTokenGroups(toks, 3, ','), items.length);
			if (at === undefined) {
				if (state.items.delete(head)) { currentSnapshot = undefined; }
			} else {
				currentSnapshot = undefined;
				items.splice(at, 0, tokenName(toks[4])!);
			}
			return;
		}
		// `c.Add 1` or `c.Add "a"`: a value, no object (issue #447).
		const literal = toks[3];
		if (head && state.items.has(head) && toks[1]?.rawText === '.' && tokenText(toks[2]) === 'add' && literal
			&& (literal.kind === 'integerLiteral' || literal.kind === 'floatLiteral' || literal.kind === 'stringLiteral')
			&& (toks[4] === undefined || toks[4].rawText === ',')) {
			const items = state.items.get(head)!;
			const at = addPosition(splitTopLevelTokenGroups(toks, 3, ','), items.length);
			if (at === undefined) {
				if (state.items.delete(head)) { currentSnapshot = undefined; }
			} else {
				currentSnapshot = undefined;
				items.splice(at, 0, HELD_VALUE);
			}
			return;
		}
		forgetChanged(toks);
	};
	const forgetChanged = (toks: readonly VbaToken[]): void => {
		for (let i = 0; i < toks.length; i++) {
			const lower = tokenName(toks[i])?.toLowerCase();
			if (!lower || (!state.classes.has(lower) && !state.items.has(lower)) || toks[i - 1]?.rawText === '.' || toks[i - 1]?.rawText === '!') {
				continue;
			}
			const next = toks[i + 1]?.rawText;
			if (next === '(') {
				continue; // an index read
			}
			if (next === '.' || next === '!') {
				// A member call leaves the object; only Count and Item leave a Collection's items.
				if (!COLLECTION_READS.has(tokenText(toks[i + 2]))) {
					if (state.items.delete(lower)) { currentSnapshot = undefined; }
				}
				continue;
			}
			forget([lower]);
		}
	};
	walkEnteringBlocks(source, proc.body, (node) => activity?.isInactive(node.span) === true, visit, {
		snapshot,
		restore: (saved: State) => {
			state.classes = new Map(saved.classes);
			state.items = new Map([...saved.items].map(([k, v]) => [k, [...v]]));
			currentSnapshot = saved;
		},
		forget: (names) => forget(names),
		touches: (stmt) => namesIn(source, stmt.span),
		enter: (node) => {
			if (state.classes.size > 0 || state.items.size > 0) {
				seen.set(node, snapshot());
			}
		},
	});
	return (node) => seen.get(node) ?? NOTHING_HELD;
}

/**
 * Where `c.Add item[, key[, before[, after]]]` puts the item among `count`,
 * as a 0-based index, or undefined when the code does not say: Before:=1 is
 * the front, After:=1 the second place, nothing the end.
 */
function addPosition(args: readonly VbaToken[][], count: number): number | undefined {
	let before: VbaToken[] | undefined;
	let after: VbaToken[] | undefined;
	for (let k = 1; k < args.length; k++) {
		const arg = args[k];
		const named = arg[1]?.rawText === ':=' ? tokenText(arg[0]) : undefined;
		const value = named ? arg.slice(2) : arg;
		const role = named ?? (k === 2 ? 'before' : k === 3 ? 'after' : 'key');
		if (value.length === 0 || role === 'key' || role === 'item') {
			continue;
		}
		if (role === 'before') {
			before = value;
		} else if (role === 'after') {
			after = value;
		} else {
			return undefined;
		}
	}
	const place = before ?? after;
	if (!place) {
		return count;
	}
	if (before && after) {
		return undefined;
	}
	const n = place.length === 1 && /^\d+$/.test(place[0].rawText) ? Number(place[0].rawText) : undefined;
	if (n === undefined || n < 1 || n > count) {
		return undefined;
	}
	return before ? n - 1 : n;
}
