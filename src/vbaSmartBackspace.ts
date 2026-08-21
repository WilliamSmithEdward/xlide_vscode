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
