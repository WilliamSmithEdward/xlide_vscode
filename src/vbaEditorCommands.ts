import * as vscode from 'vscode';

export function registerVbaEditorCommands(context: vscode.ExtensionContext): void {
	context.subscriptions.push(
		vscode.commands.registerCommand('xlide.vba.smartBackspace', async () => {
			const editor = vscode.window.activeTextEditor;
			if (!editor || editor.document.languageId !== 'vba') {
				await vscode.commands.executeCommand('deleteLeft');
				return;
			}
			const handled = await clearEmptyContinuedComment(editor);
			if (!handled) {
				await vscode.commands.executeCommand('deleteLeft');
			}
		}),
		vscode.commands.registerCommand('xlide.vba.smartTab', async () => {
			const editor = vscode.window.activeTextEditor;
			if (!editor || editor.document.languageId !== 'vba') {
				await vscode.commands.executeCommand('editor.action.indentLines');
				return;
			}
			await clearEmptyContinuedComment(editor);
			await vscode.commands.executeCommand('editor.action.indentLines');
		}),
		vscode.commands.registerCommand(
			'xlide.vba.leaveSnippetAndCursorMove',
			async (direction: CursorDirection) => {
				const move = cursorMoveFor(direction);
				if (!move) {
					return;
				}
				await vscode.commands.executeCommand('leaveSnippet');
				await vscode.commands.executeCommand('cursorMove', move);
			},
		),
	);
}

type CursorDirection = 'up' | 'down' | 'left' | 'right';

function cursorMoveFor(direction: CursorDirection): Record<string, unknown> | undefined {
	switch (direction) {
		case 'up':
		case 'down':
			return { to: direction, by: 'line', value: 1 };
		case 'left':
		case 'right':
			return { to: direction, by: 'character', value: 1 };
		default:
			return undefined;
	}
}

async function clearEmptyContinuedComment(editor: vscode.TextEditor): Promise<boolean> {
	if (editor.selections.length !== 1) {
		return false;
	}
	const selection = editor.selection;
	if (!selection.isEmpty || selection.active.line === 0) {
		return false;
	}
	const document = editor.document;
	const position = selection.active;
	const line = document.lineAt(position.line).text;
	const before = line.slice(0, position.character);
	const after = line.slice(position.character);
	if (after.trim().length > 0) {
		return false;
	}
	const match = /^(\s*)('''|') ?$/.exec(before);
	if (!match) {
		return false;
	}
	const previous = document.lineAt(position.line - 1).text.trimStart();
	if (!previous.startsWith(match[2])) {
		return false;
	}
	const markerStart = match[1].length;
	return editor.edit((edit) => {
		edit.delete(new vscode.Range(
			new vscode.Position(position.line, markerStart),
			position,
		));
	});
}
