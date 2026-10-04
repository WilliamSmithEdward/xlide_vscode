// Backspace on a blank, indented line clears the whole indent in one press.
//
// XLIDE turns off `editor.trimAutoWhitespace` so a blank line keeps the indent
// the editor gave it, the way the VBE does: press Enter twice, arrow back up,
// and the caret is still at the indent rather than at column 1 (issue #43).
//
// That option alone trades one annoyance for another. With the indent kept, a
// blank line the developer no longer wants costs one Backspace per tab stop,
// because the editor's own `useTabStops` deletes a stop at a time. This rule
// makes it one press, so indented blank lines are both normal and cheap to
// remove.
//
// Pure logic, no `vscode` dependency, so it is unit-tested directly.

/**
 * True when Backspace should clear a blank line's whole indent rather than
 * delete one tab stop.
 *
 * Only for a single caret sitting on a line that is nothing but whitespace,
 * with whitespace to its left. Any content on the line - even after the caret -
 * leaves Backspace alone, because deleting the indent would then move text.
 */
export function smartBackspaceShouldClearIndent(
	lineText: string,
	character: number,
	isEmptySelection: boolean,
): boolean {
	if (!isEmptySelection || character <= 0) {
		return false;
	}
	// A whitespace-only line, and the caret is within it. `character` can exceed
	// the text length in a virtual-space editor; the check stays true there
	// because everything to the left is still whitespace.
	return /^[ \t]*$/.test(lineText) && lineText.length > 0;
}

/** Start of an empty marker continued from the preceding comment, if any. */
export function emptyContinuedCommentMarkerStart(
    lineText: string, character: number, previousLine: string | undefined,
): number | undefined {
    if (previousLine === undefined || lineText.slice(character).trim().length > 0) { return undefined; }
    const match = /^(\s*)('+) ?$/.exec(lineText.slice(0, character));
    return match && previousLine.trimStart().startsWith(match[2]) ? match[1].length : undefined;
}

/** Remaining smart cleanup after the renderer has already deleted natively. */
export function remainingBackspaceCleanup(
    beforeLine: string, beforeCaret: number, previousLine: string | undefined,
    deletedStart: number, deletedEnd: number, afterLine: string,
): { start: number; end: number } | undefined {
    if (deletedStart < 0 || deletedStart >= deletedEnd || deletedEnd !== beforeCaret ||
        deletedEnd > beforeLine.length ||
        afterLine !== beforeLine.slice(0, deletedStart) + beforeLine.slice(deletedEnd)) { return undefined; }
    const start = smartBackspaceShouldClearIndent(beforeLine, beforeCaret, true)
        ? 0 : emptyContinuedCommentMarkerStart(beforeLine, beforeCaret, previousLine);
    return start !== undefined && start < deletedStart ? { start, end: deletedStart } : undefined;
}
