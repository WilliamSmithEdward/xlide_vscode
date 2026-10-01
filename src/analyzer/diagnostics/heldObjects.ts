// What object a local is known to hold at a statement, and the classes of
// the items a Collection local holds (issue #246). Read from `Set x = New C`,
// from `Set x = y` with y known, and from `c.Add New C`. Anything else that
// names a local whole, which may pass it ByRef or replace it, ends what is
// known of it; a member call or an index read (`x.Foo`, `c(1)`, `c.Count`)
// does not. Blocks are entered as issue #237 enters them.

import type { ConditionalActivityTracker } from '../conditional/conditionalCompilation';
import { statementLabelDeclaration } from '../flow/procedureLabels';
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

/** Members that read a Collection without changing it. */
const COLLECTION_READS: ReadonlySet<string> = new Set(['count', 'item']);

/**
 * What each statement, and each block as it is entered, sees. A statement
 * the walk never reached sees nothing.
 */
export function heldObjectsAt(
	source: string,
	proc: ProcedureNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
): (node: BodyNode) => HeldObjects {
	const seen = new Map<BodyNode, HeldObjects>();
	const state: State = { classes: new Map(), items: new Map() };
	// `Dim c As New Collection` holds an empty one from the start.
	for (const child of procedureSymbolFor(symbols, proc)?.children ?? []) {
		if (child.kind === 'localVariable' && child.isAutoInstantiated && !child.isArray && child.asType) {
			state.classes.set(child.name.toLowerCase(), child.asType);
			if (normalizeType(child.asType) === 'collection') {
				state.items.set(child.name.toLowerCase(), []);
			}
		}
	}
	const snapshot = (): HeldObjects => ({ classes: new Map(state.classes), items: new Map([...state.items].map(([k, v]) => [k, [...v]])) });
	const forget = (names: Iterable<string>): void => {
		for (const lower of names) {
			state.classes.delete(lower);
			state.items.delete(lower);
		}
	};
	const visit = (node: BodyNode): void => {
		if (!isLeafStatement(node)) {
			return;
		}
		const toks = statementTokensAfterLeadingLabel(source, node.span).filter((tok) => tok.kind !== 'comment');
		if (statementLabelDeclaration(source, node.span) || tokenText(toks[0]) === 'gosub') {
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
			const held = created ?? (from ? state.classes.get(from) : undefined);
			// The value's own holder may now change it unseen.
			forget([lower, ...namesIn(source, node.span)].filter((name) => name !== from || !held));
			if (from && held) {
				state.items.delete(from);
			}
			if (held) {
				state.classes.set(lower, held);
				if (created && normalizeType(created) === 'collection') {
					state.items.set(lower, []);
				}
			}
			return;
		}
		// `c.Add New Flat1` puts a Flat1 at the end.
		const head = tokenName(toks[0])?.toLowerCase();
		if (head && state.items.has(head) && toks[1]?.rawText === '.' && tokenText(toks[2]) === 'add'
			&& tokenText(toks[3]) === 'new' && tokenName(toks[4]) && (toks[5] === undefined || toks[5].rawText === ',')) {
			state.items.get(head)!.push(tokenName(toks[4])!);
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
					state.items.delete(lower);
				}
				continue;
			}
			forget([lower]);
		}
	};
	walkEnteringBlocks(source, proc.body, (node) => activity?.isInactive(node.span) === true, visit, {
		snapshot: () => ({ classes: new Map(state.classes), items: new Map([...state.items].map(([k, v]) => [k, [...v]])) }),
		restore: (saved: State) => {
			state.classes = new Map(saved.classes);
			state.items = new Map([...saved.items].map(([k, v]) => [k, [...v]]));
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
