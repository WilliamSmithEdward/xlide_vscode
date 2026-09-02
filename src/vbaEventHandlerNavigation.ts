// The VBE's double-click, for both designers: the control's event handler
// in the code - navigate to it, or append the stub the VBE would write and
// land the cursor inside it. Not a project write: the stub is a document
// edit the user saves like any typed code.

import * as vscode from 'vscode';

/**
 * Shows `document` at the handler `Sub <handler>(`, appending
 * `Private Sub <handler>(<params>)` with an empty body when the document
 * has none, and leaves the cursor on the line inside it.
 */
export async function openOrCreateEventHandler(document: vscode.TextDocument, handler: string, params = ''): Promise<void> {
	const escaped = handler.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
	const signature = new RegExp(`^[ \\t]*(?:Private\\s+|Public\\s+|Friend\\s+)?Sub\\s+${escaped}\\s*\\(`, 'im');
	const text = document.getText();
	if (!signature.test(text)) {
		const eol = document.eol === vscode.EndOfLine.CRLF ? '\r\n' : '\n';
		const lead = text.length === 0 ? '' : text.endsWith(eol) ? eol : eol + eol;
		const edit = new vscode.WorkspaceEdit();
		edit.insert(document.uri, document.positionAt(text.length), `${lead}Private Sub ${handler}(${params})${eol}${eol}End Sub${eol}`);
		await vscode.workspace.applyEdit(edit);
	}
	const match = signature.exec(document.getText());
	const editor = await vscode.window.showTextDocument(document, {
		viewColumn: vscode.ViewColumn.One,
		preserveFocus: false,
	});
	if (match) {
		const line = Math.min(document.positionAt(match.index).line + 1, document.lineCount - 1);
		const at = new vscode.Position(line, 0);
		editor.selection = new vscode.Selection(at, at);
		editor.revealRange(new vscode.Range(at, at), vscode.TextEditorRevealType.InCenterIfOutsideViewport);
	}
}
