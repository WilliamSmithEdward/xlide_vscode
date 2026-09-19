// The canonical-case controller's touch tracking, driven through its real
// handlers with in-memory documents and editors. Recasing is for text the
// user is typing: a line counts once the user changed it, and the line is
// recased when the user pauses or leaves it.
//
// A reload is no such change. When an agent writes a module that is open, VS
// Code reloads the document from the workbook, and the reload arrives as an
// ordinary content change. It was taken for typing, so XLIDE recased the
// agent's text 200 ms later and left the document dirty, out of step with the
// workbook - measured in a real VS Code, where a reloaded `debug.print 1` came
// back as `Debug.Print 1` and unsaved. A reload leaves the document matching
// its file, which typing never does, and that is what tells them apart.

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type * as vscodeTypes from 'vscode';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
	window: { visibleTextEditors: [], activeTextEditor: undefined },
}));

import * as vscode from 'vscode';
import { VbaCanonicalCaseController } from '../src/vbaCanonicalCaseController';

const BOOK = process.platform === 'win32' ? '/C:/Book.xlsm' : '/work/Book.xlsm';
const SOURCE = 'option explicit\nsub go()\nend sub\n';

interface FakeDocument {
	uri: { scheme: string; path: string; toString: () => string };
	languageId: string;
	isDirty: boolean;
	isClosed: boolean;
	version: number;
	lineCount: number;
	lineAt(line: number): { text: string; range: { start: vscode.Position; end: vscode.Position } };
	offsetAt(position: vscode.Position): number;
	positionAt(offset: number): vscode.Position;
	getText(range?: { start: vscode.Position; end: vscode.Position }): string;
}

interface FakeEditor {
	document: FakeDocument;
	selection: { active: vscode.Position };
	edit: ReturnType<typeof vi.fn>;
	/** The replacement texts of every edit made through this editor. */
	replaced: string[];
}

function fakeDocument(text: string): FakeDocument {
	const lines = text.split('\n');
	const starts: number[] = [];
	let at = 0;
	for (const line of lines) {
		starts.push(at);
		at += line.length + 1;
	}
	const offsetAt = (position: vscode.Position): number => starts[position.line] + position.character;
	const positionAt = (offset: number): vscode.Position => {
		let line = 0;
		while (line + 1 < starts.length && starts[line + 1] <= offset) {
			line++;
		}
		return new vscode.Position(line, offset - starts[line]);
	};
	const value = `xlide-vba:${BOOK}/Module1.bas`;
	return {
		uri: { scheme: 'xlide-vba', path: `${BOOK}/Module1.bas`, toString: () => value },
		languageId: 'vba',
		isDirty: false,
		isClosed: false,
		version: 1,
		lineCount: lines.length,
		lineAt: (line) => ({
			text: lines[line],
			range: { start: new vscode.Position(line, 0), end: new vscode.Position(line, lines[line].length) },
		}),
		offsetAt,
		positionAt,
		getText: (range) => (range ? text.slice(offsetAt(range.start), offsetAt(range.end)) : text),
	};
}

const visible = (): FakeEditor[] => (vscode.window as unknown as { visibleTextEditors: FakeEditor[] }).visibleTextEditors;

/** An editor showing the document, which VS Code refuses to edit once it is no longer visible. */
function fakeEditor(document: FakeDocument, line = 0): FakeEditor {
	const editor: FakeEditor = {
		document,
		selection: { active: new vscode.Position(line, 0) },
		replaced: [],
		edit: vi.fn(async (callback: (builder: { replace: (range: unknown, text: string) => void }) => void) => {
			if (!visible().includes(editor)) {
				throw new Error('TextEditor#edit not possible on closed editors');
			}
			callback({ replace: (_range, text) => { editor.replaced.push(text); } });
			return true;
		}),
	};
	return editor;
}

function show(editor: FakeEditor): void {
	visible().push(editor);
	(vscode.window as unknown as { activeTextEditor: FakeEditor }).activeTextEditor = editor;
}

function close(editor: FakeEditor): void {
	visible().splice(visible().indexOf(editor), 1);
	(vscode.window as unknown as { activeTextEditor?: FakeEditor }).activeTextEditor = undefined;
}

/** A content change to the document: text put in place of a stretch of one line. */
function changed(
	document: FakeDocument,
	line: number,
	from: number,
	to: number,
	text: string,
): vscodeTypes.TextDocumentChangeEvent {
	document.version += 1;
	const start = new vscode.Position(line, from);
	const end = new vscode.Position(line, to);
	return {
		document,
		contentChanges: [{ range: { start, end, isEmpty: from === to }, rangeOffset: 0, rangeLength: to - from, text }],
		reason: undefined,
	} as unknown as vscodeTypes.TextDocumentChangeEvent;
}

/** What VS Code sends when a document's dirty state moves: no content changes. */
function dirtyStateMoved(document: FakeDocument, isDirty: boolean): vscodeTypes.TextDocumentChangeEvent {
	document.isDirty = isDirty;
	return { document, contentChanges: [], reason: undefined } as unknown as vscodeTypes.TextDocumentChangeEvent;
}

function controller(): VbaCanonicalCaseController {
	return new VbaCanonicalCaseController({ cachedEditorProjectContext: () => undefined } as never);
}

beforeEach(() => {
	vi.useFakeTimers();
	visible().length = 0;
	(vscode.window as unknown as { activeTextEditor?: FakeEditor }).activeTextEditor = undefined;
});

afterEach(() => {
	vi.useRealTimers();
});

describe('what the canonical-case controller counts as typing', () => {
	it('leaves a reloaded line alone: the change left the document matching its file', async () => {
		const document = fakeDocument(SOURCE);
		const editor = fakeEditor(document);
		show(editor);
		const casing = controller();

		// The reload replaces line 0 and the document stays clean.
		casing.handleTextDocumentChange(changed(document, 0, 0, 15, 'option explicit'));
		await vi.advanceTimersByTimeAsync(1000);

		expect(editor.edit).not.toHaveBeenCalled();
	});

	it('recases the line the user typed on, once the document has turned dirty', async () => {
		const document = fakeDocument(SOURCE);
		const editor = fakeEditor(document);
		show(editor);
		const casing = controller();

		// A keystroke on a clean document: VS Code reports it clean at the
		// change and dirty just after, in an event of its own.
		casing.handleTextDocumentChange(changed(document, 0, 15, 15, ' '));
		casing.handleTextDocumentChange(dirtyStateMoved(document, true));
		await vi.advanceTimersByTimeAsync(1000);

		expect(editor.edit).toHaveBeenCalledTimes(1);
		expect(editor.replaced).toEqual(['Option', 'Explicit']);
	});

	it('recases at once on Enter in a document the user is already editing', async () => {
		const document = fakeDocument(SOURCE);
		document.isDirty = true;
		const editor = fakeEditor(document);
		show(editor);
		const casing = controller();

		casing.handleTextDocumentChange(changed(document, 0, 15, 15, '\n'));
		await vi.advanceTimersByTimeAsync(0);

		expect(editor.replaced).toEqual(['Option', 'Explicit']);
	});

	it('recases the first Enter after a save even when the dirty state is reported late', async () => {
		// Enter on a document the user is editing recases at once. On a clean
		// one the change could still be a reload, so it waits for the idle
		// pass - which also means a dirty report that comes after the moment
		// the immediate pass would have run cannot lose the recase.
		const document = fakeDocument(SOURCE);
		const editor = fakeEditor(document);
		show(editor);
		const casing = controller();

		casing.handleTextDocumentChange(changed(document, 0, 15, 15, '\n'));
		await vi.advanceTimersByTimeAsync(0);
		casing.handleTextDocumentChange(dirtyStateMoved(document, true));
		await vi.advanceTimersByTimeAsync(1000);

		expect(editor.replaced).toEqual(['Option', 'Explicit']);
	});

	it('forgets a line a reload touched, so leaving it after a later edit does not recase it', async () => {
		const document = fakeDocument(SOURCE);
		const editor = fakeEditor(document);
		show(editor);
		const casing = controller();
		casing.handleTextDocumentChange(changed(document, 0, 0, 15, 'option explicit'));
		await vi.advanceTimersByTimeAsync(1000);

		// The user edits elsewhere later, then the caret leaves line 0.
		document.isDirty = true;
		casing.handleActiveEditorChange(editor as unknown as vscodeTypes.TextEditor);
		editor.selection = { active: new vscode.Position(2, 0) };
		casing.handleSelectionChange({ textEditor: editor } as unknown as vscodeTypes.TextEditorSelectionChangeEvent);
		await vi.advanceTimersByTimeAsync(1000);

		expect(editor.edit).not.toHaveBeenCalled();
	});

	it('writes the casing of a line typed just before a save into the save', async () => {
		// Auto-save with a short delay saved before the pause, the pause then
		// found a document matching its file, and `long` was never recased.
		const document = fakeDocument(SOURCE);
		const editor = fakeEditor(document);
		show(editor);
		const casing = controller();
		casing.handleTextDocumentChange(changed(document, 0, 15, 15, ' '));
		casing.handleTextDocumentChange(dirtyStateMoved(document, true));

		const edits = casing.pendingEditsForSave(document as never);
		casing.handleTextDocumentChange(dirtyStateMoved(document, false));
		await vi.advanceTimersByTimeAsync(1000);

		expect(edits.map((edit) => edit.newText)).toEqual(['Option', 'Explicit']);
		// The save took the casing, so nothing is left to make it dirty again.
		expect(editor.edit).not.toHaveBeenCalled();
	});

	it('writes nothing into a save for a line only a reload touched', async () => {
		const document = fakeDocument(SOURCE);
		const editor = fakeEditor(document);
		show(editor);
		const casing = controller();
		casing.handleTextDocumentChange(changed(document, 0, 0, 15, 'option explicit'));

		expect(casing.pendingEditsForSave(document as never)).toEqual([]);
	});

	it('does not recase after the document was saved', async () => {
		const document = fakeDocument(SOURCE);
		document.isDirty = true;
		const editor = fakeEditor(document);
		show(editor);
		const casing = controller();
		casing.handleActiveEditorChange(editor as unknown as vscodeTypes.TextEditor);
		casing.handleTextDocumentChange(changed(document, 0, 15, 15, ' '));

		// Saved before the pause was over: the document matches its file again.
		casing.handleTextDocumentChange(dirtyStateMoved(document, false));
		await vi.advanceTimersByTimeAsync(1000);

		expect(editor.edit).not.toHaveBeenCalled();
	});
});

describe('an editor that closes under a pending recase', () => {
	it('is not edited once it has closed', async () => {
		const document = fakeDocument(SOURCE);
		document.isDirty = true;
		const editor = fakeEditor(document);
		show(editor);
		const casing = controller();

		casing.handleTextDocumentChange(changed(document, 0, 15, 15, ' '));
		close(editor);
		await vi.advanceTimersByTimeAsync(1000);

		expect(editor.edit).not.toHaveBeenCalled();
	});

	it('edits through another editor still showing the document', async () => {
		const document = fakeDocument(SOURCE);
		document.isDirty = true;
		const first = fakeEditor(document);
		const second = fakeEditor(document);
		show(second);
		show(first);
		const casing = controller();

		casing.handleTextDocumentChange(changed(document, 0, 15, 15, ' '));
		close(first);
		await vi.advanceTimersByTimeAsync(1000);

		expect(first.edit).not.toHaveBeenCalled();
		expect(second.replaced).toEqual(['Option', 'Explicit']);
	});

	it('drops the edit when the editor closes while it is applied, and throws anything else', async () => {
		const document = fakeDocument(SOURCE);
		document.isDirty = true;
		const editor = fakeEditor(document);
		show(editor);
		const casing = controller();

		editor.edit.mockImplementationOnce(async () => {
			close(editor);
			throw new Error('Illegal argument: TextEditor(vs.editor.ICodeEditor:2,$model3)');
		});
		await expect(casing.applyCanonicalCaseForLine(document as never, 0, editor as never)).resolves.toBeUndefined();

		show(editor);
		editor.edit.mockImplementationOnce(async () => {
			throw new Error('the edit failed for another reason');
		});
		await expect(casing.applyCanonicalCaseForLine(document as never, 0, editor as never))
			.rejects.toThrow('the edit failed for another reason');
	});
});
