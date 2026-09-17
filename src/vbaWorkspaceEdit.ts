import * as vscode from 'vscode';

/**
 * The analyzer's offset-based edits as one workspace edit against an open
 * document. Quick fixes, refactorings and the analysis results view all land
 * their edits this way.
 */
export function workspaceEditFor(
    document: vscode.TextDocument,
    edits: readonly { span: { start: number; end: number }; newText: string }[],
): vscode.WorkspaceEdit {
    const edit = new vscode.WorkspaceEdit();
    for (const textEdit of edits) {
        edit.replace(
            document.uri,
            new vscode.Range(
                document.positionAt(textEdit.span.start),
                document.positionAt(textEdit.span.end),
            ),
            textEdit.newText,
        );
    }
    return edit;
}
