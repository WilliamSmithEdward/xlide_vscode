// Pure decision logic for the VBA Smart Tab command. Kept free of any vscode
// dependency so it can be unit-tested directly.

/**
 * Decides whether Tab in a VBA editor should indent the whole line/block (VBE-like)
 * rather than insert a tab at the caret. Indent on a multi-line selection, or when
 * the caret is on a blank line or within the leading whitespace; otherwise (caret
 * inside the line content, e.g. at the end of a line) a plain tab is inserted.
 */
export function smartTabShouldIndentLine(
    lineText: string,
    caretColumn: number,
    selectionIsEmpty: boolean,
    selectionSpansLines: boolean,
): boolean {
    if (selectionSpansLines) {
        return true;
    }
    if (!selectionIsEmpty) {
        return false;
    }
    const firstNonWs = lineText.search(/\S/);
    return firstNonWs === -1 || caretColumn <= firstNonWs;
}
