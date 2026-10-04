// What a module's code may do to its workbook's sheets while it runs (issue
// #229). The sheets a workbook was saved with say which names and indexes
// `ThisWorkbook.Sheets(...)` can reach, but only until code adds, copies or
// renames one. This scan finds those operations in the text, so the check
// judges a sheet only when no code in the project could have made it.
//
// It errs towards finding too much: a With block's `.Add` counts as adding a
// sheet whatever the With names, and so does a `.Copy` given an argument that
// no range word explains.

import { tokenizeCached } from '../lexer/tokenize';
import type { VbaToken } from '../lexer/tokenKinds';

/** A sheet of the saved workbook, as the host read it from the file. */
export interface WorkbookSheetInfo {
	name: string;
	kind: 'worksheet' | 'chartsheet' | 'dialogsheet' | 'macrosheet';
}

export interface SheetChanges {
	/** Code adds or copies a sheet, so the count and the default names can grow. */
	addsSheets: boolean;
	/** Lowercased names that code assigns to some `.Name` as a literal. */
	namesAssigned: ReadonlySet<string>;
	/** Code assigns some `.Name` a value that is not a literal, so any name can appear. */
	assignsComputedName: boolean;
}

const SHEET_COLLECTIONS = new Set(['sheets', 'worksheets', 'charts']);
const RANGE_WORDS = new Set(['range', 'cells', 'rows', 'columns', 'usedrange', 'selection', 'destination', 'currentregion', 'entirerow', 'entirecolumn', 'offset', 'resize']);
/** Tokens after which a member chain begins a statement rather than an expression. */
const STATEMENT_OPENERS = new Set(['then', 'else']);

/** The sheet operations in one module's code. */
export function sheetChangesIn(source: string): SheetChanges {
	const names = new Set<string>();
	let addsSheets = false;
	let assignsComputedName = false;
	const statement: VbaToken[] = [];
	const flush = (): void => {
		const found = scanStatement(statement, names);
		addsSheets ||= found.addsSheets;
		assignsComputedName ||= found.assignsComputedName;
		statement.length = 0;
	};
	for (const token of tokenizeCached(source)) {
		if (token.kind === 'comment') {
			continue;
		}
		if (token.kind === 'newline' || token.kind === 'colon') {
			flush();
			continue;
		}
		statement.push(token);
	}
	flush();
	return { addsSheets, namesAssigned: names, assignsComputedName };
}

/** The union of several modules' sheet operations. */
export function mergeSheetChanges(all: Iterable<SheetChanges>): SheetChanges {
	const names = new Set<string>();
	let addsSheets = false;
	let assignsComputedName = false;
	for (const changes of all) {
		addsSheets ||= changes.addsSheets;
		assignsComputedName ||= changes.assignsComputedName;
		for (const name of changes.namesAssigned) {
			names.add(name);
		}
	}
	return { addsSheets, namesAssigned: names, assignsComputedName };
}

function scanStatement(toks: readonly VbaToken[], names: Set<string>): { addsSheets: boolean; assignsComputedName: boolean } {
	let addsSheets = false;
	let assignsComputedName = false;
	const lower = (i: number): string => toks[i]?.rawText.toLowerCase() ?? '';
	// Every Copy member asks the same statement-wide range-word question.
	let hasRangeWords: boolean | undefined;
	for (let i = 0; i < toks.length; i++) {
		if (toks[i].rawText !== '.') {
			continue;
		}
		const member = lower(i + 1);
		if (member === 'add' || member === 'add2') {
			// `Worksheets.Add`, or `.Add` inside a With whose object is not in view.
			const before = i > 0 ? toks[i - 1] : undefined;
			if (!before || before.kind === 'keyword' || SHEET_COLLECTIONS.has(before.rawText.toLowerCase())) {
				addsSheets = true;
			}
		} else if (member === 'copy' && i + 2 < toks.length && !(hasRangeWords ??= toks.some((t) => RANGE_WORDS.has(t.rawText.toLowerCase())))) {
			// `ws.Copy After:=...` makes a sheet; `ws.Copy` alone makes a workbook.
			addsSheets = true;
		} else if (member === 'name' && toks[i + 2]?.rawText === '=') {
			const value = toks[i + 3];
			// A lone literal, maybe followed by a single-line If's Else.
			if (value?.kind === 'stringLiteral' && (i + 4 === toks.length || STATEMENT_OPENERS.has(lower(i + 4)))) {
				names.add(value.rawText.slice(1, -1).replace(/""/g, '"').toLowerCase());
			} else if (startsStatement(toks, i)) {
				assignsComputedName = true;
			}
		}
	}
	return { addsSheets, assignsComputedName };
}

/**
 * True when the member chain ending at the `.` at `dot` starts the
 * statement, as an assignment target does, rather than sitting inside a
 * condition such as `If ws.Name = s Then`.
 */
function startsStatement(toks: readonly VbaToken[], dot: number): boolean {
	let i = dot - 1;
	while (i >= 0) {
		const text = toks[i].rawText;
		if (text === ')') {
			let depth = 0;
			for (; i >= 0; i--) {
				if (toks[i].rawText === ')') { depth++; }
				if (toks[i].rawText === '(' && --depth === 0) { break; }
			}
			i--;
			continue;
		}
		// `Me.Name = s` in a sheet's own module renames that sheet.
		if (toks[i].kind === 'identifier' || toks[i].kind === 'bracketedIdentifier' || text === '.' || text === '!' || text.toLowerCase() === 'me') {
			i--;
			continue;
		}
		break;
	}
	return i < 0 || STATEMENT_OPENERS.has(toks[i].rawText.toLowerCase());
}
