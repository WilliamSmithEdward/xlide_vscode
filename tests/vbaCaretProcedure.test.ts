// The one answer to "which procedure is the caret in", shared by the status
// bar and the explorer (github.com/WilliamSmithEdward/xlide_vscode/issues/66).
// The point of the tracker is what it does NOT do: the caret moves on every
// keystroke and arrow key, and almost none of those moves leave the procedure.

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

const mock = vi.hoisted(() => ({
    editorChanged: undefined as unknown,
    selectionChanged: undefined as unknown,
}));

vi.mock('vscode', async () => {
    const base = await import('./helpers/vscodeMock');
    mock.editorChanged = new base.EventEmitter<unknown>();
    mock.selectionChanged = new base.EventEmitter<unknown>();
    return base.vscodeMock({
        window: {
            activeTextEditor: undefined,
            onDidChangeActiveTextEditor: (mock.editorChanged as { event: unknown }).event,
            onDidChangeTextEditorSelection: (mock.selectionChanged as { event: unknown }).event,
        },
    });
});

// The URI codec has its own tests; this one is about the tracker, so the
// document-to-module answer is stubbed rather than encoded.
vi.mock('../src/vbaDocumentLocation', () => ({
    moduleLocationOfDocument: (document: { uri: { scheme: string } }) =>
        (document.uri.scheme === 'xlide-vba'
            ? { projectPath: 'C:\\work\\Book.xlsm', moduleName: 'Helpers', native: false }
            : undefined),
}));

import * as vscode from 'vscode';
import { VbaCaretProcedureTracker, type VbaCaretPosition } from '../src/vbaCaretProcedure';

const SOURCE = [
    /* 0 */ 'Option Explicit',
    /* 1 */ '',
    /* 2 */ 'Sub Post()',
    /* 3 */ '    Debug.Print 1',
    /* 4 */ '    Debug.Print 2',
    /* 5 */ 'End Sub',
    /* 6 */ '',
    /* 7 */ 'Function Total() As Long',
    /* 8 */ 'End Function',
].join('\n');

/** An editor over a project module's virtual document. */
function editor(line: number, text = SOURCE, version = 1) {
    return {
        document: {
            uri: { scheme: 'xlide-vba', toString: () => 'xlide-vba:/Book.xlsm/Helpers.bas' },
            version,
            getText: vi.fn(() => text),
        },
        selection: { active: { line } },
    };
}

/** The active editor is a plain property on the stub, so tests set it. */
function setActiveEditor(value: unknown): void {
    (vscode.window as unknown as { activeTextEditor: unknown }).activeTextEditor = value;
}

function fireSelection(): void {
    (mock.selectionChanged as { fire: (v: unknown) => void })
        .fire({ textEditor: vscode.window.activeTextEditor });
}

/** Move the caret, optionally onto a newly edited version of the module. */
function moveTo(line: number, text?: string, version?: number): void {
    setActiveEditor(editor(line, text, version));
    fireSelection();
}

let tracker: VbaCaretProcedureTracker | undefined;
let seen: Array<VbaCaretPosition | undefined> = [];

/** Start tracking from wherever the caret has been put. */
function track(): VbaCaretProcedureTracker {
    tracker = new VbaCaretProcedureTracker();
    seen = [];
    tracker.onDidChange((position) => seen.push(position));
    return tracker;
}

beforeEach(() => setActiveEditor(undefined));

// The stub's emitters outlive a test, so a tracker left listening would go on
// reporting into the next one's expectations.
afterEach(() => {
    tracker?.dispose();
    tracker = undefined;
});

describe('what the tracker reports', () => {
    it('names the module and the procedure the caret sits in', () => {
        setActiveEditor(editor(3));
        expect(track().current).toMatchObject({
            projectPath: 'C:\\work\\Book.xlsm',
            moduleName: 'Helpers',
            native: false,
            label: 'Sub Post',
        });
    });

    it("says (Declarations) above the module's first procedure", () => {
        setActiveEditor(editor(0));
        expect(track().current?.label).toBe('(Declarations)');
    });

    it('reports nothing with no editor at all', () => {
        expect(track().current).toBeUndefined();
    });

    it('reports nothing for a document no project claims', () => {
        setActiveEditor({
            document: { uri: { scheme: 'file' }, version: 1, getText: vi.fn(() => '') },
            selection: { active: { line: 0 } },
        });
        expect(track().current).toBeUndefined();
    });
});

describe('when it fires', () => {
    it('fires on the move that crosses into another procedure', () => {
        setActiveEditor(editor(3));
        track();
        moveTo(8);
        expect(seen.map((p) => p?.label)).toEqual(['Function Total']);
    });

    it('stays quiet while the caret moves inside one procedure', () => {
        setActiveEditor(editor(3));
        track();
        moveTo(4);
        moveTo(5);
        moveTo(2);
        expect(seen).toEqual([]);
    });

    it('fires when the editor leaves every module', () => {
        setActiveEditor(editor(3));
        const started = track();
        setActiveEditor(undefined);
        (mock.editorChanged as { fire: (v: unknown) => void }).fire(undefined);
        expect(seen).toEqual([undefined]);
        expect(started.current).toBeUndefined();
    });
});

describe('what it costs', () => {
    it('reads the module once, however far the caret moves', () => {
        const one = editor(3);
        setActiveEditor(one);
        track();
        for (const line of [4, 5, 8, 0, 3]) {
            one.selection.active.line = line;
            fireSelection();
        }
        expect(one.document.getText).toHaveBeenCalledTimes(1);
    });

    it('reads it again once the text has actually changed', () => {
        setActiveEditor(editor(3));
        track();
        // A new version with the procedure renamed: same caret line, new answer.
        moveTo(3, SOURCE.replace('Sub Post()', 'Sub Posted()'), 2);
        expect(seen.map((p) => p?.label)).toEqual(['Sub Posted']);
    });
});
