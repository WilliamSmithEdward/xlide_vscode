import * as vscode from 'vscode';
import { isVbaDocument } from './xlideFileSystem';
import { smartTabShouldIndentLine } from './vbaSmartTab';
import { emptyContinuedCommentMarkerStart, remainingBackspaceCleanup, smartBackspaceShouldClearIndent } from './vbaSmartBackspace';

export const BACKSPACE_NEEDS_EXTENSION_CONTEXT = 'xlide.vba.backspaceNeedsExtension';

/** Ordinary code deletion must stay native when the extension host is busy. */
export function backspaceNeedsExtension(editor: vscode.TextEditor | undefined): boolean {
	if (!editor || editor.document.isClosed || !isVbaDocument(editor.document) || editor.selections.length !== 1) { return false; }
	const selection = editor.selection;
	if (!selection.isEmpty || selection.active.line < 0 || selection.active.line >= editor.document.lineCount) { return false; }
	return smartBackspaceShouldClearIndent(editor.document.lineAt(selection.active.line).text,
		selection.active.character, true) || emptyContinuedCommentStart(editor) !== undefined;
}

interface BackspaceState {
    editor: vscode.TextEditor;
    document: vscode.TextDocument;
    version: number;
    line: number;
    character: number;
    text: string;
    previousLine?: string;
}
interface PendingBackspaceCleanup {
    before: BackspaceState;
    version: number;
    character: number;
    range: vscode.Range;
    requested: boolean;
}

function backspaceState(): BackspaceState | undefined {
    const editor = vscode.window.activeTextEditor;
    if (!editor || editor.document.isClosed || !isVbaDocument(editor.document) || editor.selections.length !== 1 ||
        !editor.selection.isEmpty) { return undefined; }
    const document = editor.document;
    const { line, character } = editor.selection.active;
    if (line < 0 || line >= document.lineCount) { return undefined; }
    const text = document.lineAt(line).text;
    if (character < 0 || character > text.length) { return undefined; }
    return { editor, document, version: document.version, line, character, text,
        previousLine: line > 0 && /^\s*'/.test(text) ? document.lineAt(line - 1).text : undefined };
}

function registerBackspaceContext(context: vscode.ExtensionContext): void {
    let previous: boolean | undefined;
    let state: BackspaceState | undefined;
    const recentStates: BackspaceState[] = [];
    let pending: PendingBackspaceCleanup | undefined;
    const update = (): void => {
        state = backspaceState();
        if (state) {
            recentStates.push(state);
            if (recentStates.length > 2) { recentStates.shift(); }
        }
        const next = backspaceNeedsExtension(vscode.window.activeTextEditor);
        if (next === previous) { return; }
        previous = next;
        void vscode.commands.executeCommand('setContext', BACKSPACE_NEEDS_EXTENSION_CONTEXT, next);
    };
    const finish = (): Thenable<unknown> | undefined => {
        const cleanup = pending;
        if (!cleanup?.requested) { return; }
        const { editor, document, line, character: beforeCharacter } = cleanup.before;
        if (vscode.window.activeTextEditor !== editor || editor.document !== document || document.isClosed ||
            document.version !== cleanup.version || editor.selections.length !== 1 || !editor.selection.isEmpty) {
            pending = undefined; return;
        }
        const caret = editor.selection.active;
        if (caret.line !== line || caret.character !== cleanup.character) {
            // The document event can precede the native caret update. Its
            // selection event will finish cleanup without a response-path wait.
            if (caret.line === line && caret.character === beforeCharacter) { return; }
            pending = undefined; return;
        }
        pending = undefined;
        // Native deletion already happened. Rejection, close or a later edit
        // must never produce a fallback deletion at a newer caret.
        return editor.edit(edit => edit.delete(cleanup.range), { undoStopBefore: false, undoStopAfter: true })
            .then(undefined, () => undefined); // A close/version race can reject the edit after dispatch.
    };
    context.subscriptions.push(
        vscode.commands.registerCommand('xlide.vba.finishBackspaceCleanup', () => {
            if (pending) { pending.requested = true; }
            return finish();
        }),
        vscode.window.onDidChangeActiveTextEditor(() => { pending = undefined; recentStates.length = 0; update(); }),
        vscode.window.onDidChangeTextEditorSelection(event => {
            if (event.textEditor !== vscode.window.activeTextEditor) { return; }
            if (pending && (event.textEditor.selection.active.line !== pending.before.line ||
                event.textEditor.selection.active.character !== pending.character)) { pending = undefined; }
            update(); void finish();
        }),
        vscode.workspace.onDidChangeTextDocument(event => {
            if (event.document !== vscode.window.activeTextEditor?.document) { return; }
            // Dirty-state notifications have no text changes and may follow the
            // native deletion before its caret notification. Preserve cleanup.
            if (!event.contentChanges?.length) { update(); return; }
            pending = undefined;
            const change = event.contentChanges?.length === 1 ? event.contentChanges[0] : undefined;
            // Native selection and document notifications can arrive in either
            // order. Retain the preceding caret alongside the latest one.
            const before = change && [...recentStates].reverse().find(snapshot =>
                snapshot.document === event.document && snapshot.version === event.document.version - 1 &&
                snapshot.line === change.range.end.line && snapshot.character === change.range.end.character);
            if (before && event.reason === undefined && before.document === event.document &&
                before.editor === vscode.window.activeTextEditor && event.document.version === before.version + 1 &&
                change?.text === '' && change.rangeLength > 0 && change.range.start.line === before.line &&
                change.range.end.line === before.line) {
                const range = remainingBackspaceCleanup(before.text, before.character, before.previousLine,
                    change.range.start.character, change.range.end.character, event.document.lineAt(before.line).text);
                if (range) { pending = { before, version: event.document.version, character: change.range.start.character,
                    range: new vscode.Range(new vscode.Position(before.line, range.start), new vscode.Position(before.line, range.end)), requested: false }; }
            }
            update();
        }),
        vscode.workspace.onDidCloseTextDocument(document => {
            if (document === state?.document || document === pending?.before.document) { state = undefined; pending = undefined; recentStates.length = 0; }
            if (document === vscode.window.activeTextEditor?.document) { update(); }
        }),
        { dispose: () => { state = undefined; pending = undefined; recentStates.length = 0; void vscode.commands.executeCommand('setContext', BACKSPACE_NEEDS_EXTENSION_CONTEXT, false); } },
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
    const { line, character } = selection.active;
    return emptyContinuedCommentMarkerStart(editor.document.lineAt(line).text, character,
        editor.document.lineAt(line - 1).text);
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
