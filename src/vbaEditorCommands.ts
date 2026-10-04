import * as vscode from 'vscode';
import { isVbaDocument } from './xlideFileSystem';
import { smartTabShouldIndentLine } from './vbaSmartTab';
import { smartBackspaceShouldClearIndent } from './vbaSmartBackspace';

export const BACKSPACE_NEEDS_EXTENSION_CONTEXT = 'xlide.vba.backspaceNeedsExtension';

/** Ordinary code deletion must stay native when the extension host is busy. */
export function backspaceNeedsExtension(editor: vscode.TextEditor | undefined): boolean {
	if (!editor || editor.document.isClosed || !isVbaDocument(editor.document) || editor.selections.length !== 1) { return false; }
	const selection = editor.selection;
	if (!selection.isEmpty || selection.active.line < 0 || selection.active.line >= editor.document.lineCount) { return false; }
	return smartBackspaceShouldClearIndent(editor.document.lineAt(selection.active.line).text,
		selection.active.character, true) || emptyContinuedCommentStart(editor) !== undefined;
}

function registerBackspaceContext(context: vscode.ExtensionContext): void {
	let previous: boolean | undefined;
	const update = (): void => {
		const next = backspaceNeedsExtension(vscode.window.activeTextEditor);
		if (next === previous) { return; }
		previous = next;
		void vscode.commands.executeCommand('setContext', BACKSPACE_NEEDS_EXTENSION_CONTEXT, next);
	};
	context.subscriptions.push(
		vscode.window.onDidChangeActiveTextEditor(update),
		vscode.window.onDidChangeTextEditorSelection(event => {
			if (event.textEditor === vscode.window.activeTextEditor) { update(); }
		}),
		vscode.workspace.onDidChangeTextDocument(event => {
			if (event.document === vscode.window.activeTextEditor?.document) { update(); }
		}),
		vscode.workspace.onDidCloseTextDocument(document => {
			if (document === vscode.window.activeTextEditor?.document) { update(); }
		}),
		{ dispose: () => { void vscode.commands.executeCommand('setContext', BACKSPACE_NEEDS_EXTENSION_CONTEXT, false); } },
	);
	update();
}

export function registerVbaEditorCommands(context: vscode.ExtensionContext): void {
	registerBackspaceContext(context);
	context.subscriptions.push(
		vscode.commands.registerCommand('xlide.vba.smartBackspace', async () => {
			const editor = vscode.window.activeTextEditor;
			if (!editor || !isVbaDocument(editor.document)) {
				await vscode.commands.executeCommand('deleteLeft');
				return;
			}
			const clearing = clearEmptyContinuedComment(editor);
			const handled = typeof clearing === 'boolean' || clearing === undefined ? clearing : await clearing;
			if (handled === undefined || handled) {
				return;
			}
			// A blank line keeps its indent here (trimAutoWhitespace is off), so
			// removing one costs a press per tab stop unless the whole indent goes
			// at once (issue #43).
			const selection = editor.selection;
			if (
				editor.selections.length === 1
				&& smartBackspaceShouldClearIndent(
					editor.document.lineAt(selection.active.line).text,
					selection.active.character,
					selection.isEmpty,
				)
			) {
				await vscode.commands.executeCommand('deleteAllLeft');
				return;
			}
			await vscode.commands.executeCommand('deleteLeft');
		}),
		vscode.commands.registerCommand('xlide.vba.smartTab', async () => {
			const editor = vscode.window.activeTextEditor;
			if (!editor || !isVbaDocument(editor.document)) {
				await vscode.commands.executeCommand('tab');
				return;
			}
			const clearing = clearEmptyContinuedComment(editor);
			const cleared = typeof clearing === 'boolean' || clearing === undefined ? clearing : await clearing;
			if (cleared === undefined) { return; }
			const selection = editor.selection;
			const lineText = editor.document.lineAt(selection.active.line).text;
			const spansLines = editor.selections.some((s) => s.start.line !== s.end.line);
			if (smartTabShouldIndentLine(lineText, selection.active.character, selection.isEmpty, spansLines)) {
				await vscode.commands.executeCommand('editor.action.indentLines');
			} else {
				// Caret inside line content (e.g. end of a line): insert a tab at the
				// cursor like a normal editor instead of shifting the whole line.
				await vscode.commands.executeCommand('tab');
			}
		}),
		vscode.commands.registerCommand(
			'xlide.vba.leaveSnippetAndCursorMove',
			async (direction: CursorDirection) => {
				const move = cursorMoveFor(direction);
				if (!move) {
					return;
				}
				const editor = vscode.window.activeTextEditor;
				const document = editor?.document;
				const version = document?.version;
				await vscode.commands.executeCommand('leaveSnippet');
				if (vscode.window.activeTextEditor !== editor || editor?.document !== document ||
					document?.isClosed || document?.version !== version) { return; }
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

function emptyContinuedCommentStart(editor: vscode.TextEditor): number | undefined {
	if (editor.document.isClosed) { return undefined; }
	if (editor.selections.length !== 1) {
		return undefined;
	}
	const selection = editor.selection;
	if (!selection.isEmpty || selection.active.line === 0) {
		return undefined;
	}
	const document = editor.document;
	const position = selection.active;
	const line = document.lineAt(position.line).text;
	const before = line.slice(0, position.character);
	const after = line.slice(position.character);
	if (after.trim().length > 0) {
		return undefined;
	}
	// Match any apostrophe run, mirroring commentContinuationText's ('+) capture,
	// so Smart Backspace clears 2- and 4+-apostrophe continued comments too.
	const match = /^(\s*)('+) ?$/.exec(before);
	if (!match) {
		return undefined;
	}
	const previous = document.lineAt(position.line - 1).text.trimStart();
	if (!previous.startsWith(match[2])) {
		return undefined;
	}
	return match[1].length;
}

function clearEmptyContinuedComment(editor: vscode.TextEditor): boolean | undefined | Thenable<boolean | undefined> {
	if (editor.document.isClosed) { return undefined; }
	const markerStart = emptyContinuedCommentStart(editor);
	if (markerStart === undefined) { return false; }
	const selection = editor.selection;
	const document = editor.document;
	const position = selection.active;
	const version = document.version;
	return editor.edit((edit) => {
		edit.delete(new vscode.Range(
			new vscode.Position(position.line, markerStart),
			position,
		));
	}).then(applied => {
		// A rejected edit must not turn into a deletion/tab at a newer caret
		// or in an editor selected while the original edit was pending.
		if (vscode.window.activeTextEditor !== editor || editor.document !== document || document.isClosed ||
			document.version < version || document.version > version + (applied ? 1 : 0)) { return undefined; }
		const current = editor.selection;
		if (!applied && (current.active.line !== selection.active.line || current.active.character !== selection.active.character ||
			current.anchor.line !== selection.anchor.line || current.anchor.character !== selection.anchor.character)) {
			return undefined;
		}
		return applied;
	});
}
