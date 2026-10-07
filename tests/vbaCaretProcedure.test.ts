// The one answer to "which procedure is the caret in", shared by the status
// bar and the explorer (github.com/WilliamSmithEdward/xlide_vscode/issues/66).
// The point of the tracker is what it does NOT do: the caret moves on every
// keystroke and arrow key, and almost none of those moves leave the procedure.

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

const mock = vi.hoisted(() => ({
    editorChanged: undefined as unknown,
    selectionChanged: undefined as unknown,
    documentChanged: undefined as unknown,
}));

vi.mock('vscode', async () => {
    const base = await import('./helpers/vscodeMock');
    mock.editorChanged = new base.EventEmitter<unknown>();
    mock.selectionChanged = new base.EventEmitter<unknown>();
    mock.documentChanged = new base.EventEmitter<unknown>();
    return base.vscodeMock({
        workspace: { onDidChangeTextDocument: (mock.documentChanged as { event: unknown }).event },
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
            lineCount: text.split('\n').length,
            lineAt: vi.fn((line: number) => ({ text: text.split('\n')[line] })),
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

function switchTo(value: ReturnType<typeof editor> | undefined): void {
    setActiveEditor(value);
    (mock.editorChanged as { fire: (v: unknown) => void }).fire(value);
}

describe('procedure ranges across tab switches', () => {
    it('scans each unchanged document only once while alternating tabs', () => {
        const a = editor(3);
        const b = editor(8);
        b.document.uri.toString = () => 'xlide-vba:/Book.xlsm/Other.bas';
        setActiveEditor(a);
        track();
        for (let i = 0; i < 100; i++) {
            switchTo(b);
            expect(tracker?.current?.label).toBe('Function Total');
            switchTo(a);
            expect(tracker?.current?.label).toBe('Sub Post');
        }
        expect(a.document.getText).toHaveBeenCalledTimes(1);
        expect(b.document.getText).toHaveBeenCalledTimes(1);
    });

    it('rescans an edited inactive document and keeps the other cache entry', () => {
        const a = editor(3);
        const b = editor(8);
        b.document.uri.toString = () => 'xlide-vba:/Book.xlsm/Other.bas';
        setActiveEditor(a);
        track();
        switchTo(b);
        a.document.version++;
        a.document.getText.mockReturnValue(SOURCE.replace('Sub Post()', 'Sub Posted()'));
        switchTo(a);
        expect(tracker?.current?.label).toBe('Sub Posted');
        switchTo(b);
        switchTo(a);
        expect(a.document.getText).toHaveBeenCalledTimes(2);
        expect(b.document.getText).toHaveBeenCalledTimes(1);
    });

    it('rescans a reopened document even when its URI and version repeat', () => {
        setActiveEditor(editor(3));
        track();
        const reopened = editor(3, SOURCE.replace('Sub Post()', 'Sub Reopened()'));
        switchTo(reopened);
        expect(tracker?.current?.label).toBe('Sub Reopened');
        expect(reopened.document.getText).toHaveBeenCalledTimes(1);
    });

    it('keeps cached ranges when focus leaves the editor and returns', () => {
        const one = editor(3);
        setActiveEditor(one);
        track();
        switchTo(undefined);
        expect(tracker?.current).toBeUndefined();
        switchTo(one);
        expect(tracker?.current?.label).toBe('Sub Post');
        expect(one.document.getText).toHaveBeenCalledTimes(1);
    });
});

describe('ordinary typing reuses procedure ranges', () => {
    function change(one: ReturnType<typeof editor>, text: string, line: number, inserted: string, endLine = line,
        notifySelection = true) {
        one.document.version++;
        one.document.lineCount = text.split('\n').length;
        one.document.getText.mockReturnValue(text);
        one.document.lineAt.mockImplementation((index) => ({ text: text.split('\n')[index] }));
        (mock.documentChanged as { fire: (v: unknown) => void }).fire({
            document: one.document,
            contentChanges: [{ text: inserted, range: { start: { line }, end: { line: endLine } } }],
        });
        if (notifySelection) { fireSelection(); }
    }

    it('reports a renamed procedure without waiting for the caret to move', () => {
        const one = editor(3);
        setActiveEditor(one);
        track();
        change(one, SOURCE.replace('Sub Post()', 'Sub Posted()'), 2, 'ed', 2, false);
        expect(tracker?.current?.label).toBe('Sub Posted');
        expect(seen.map(position => position?.label)).toEqual(['Sub Posted']);
        fireSelection();
        expect(one.document.getText).toHaveBeenCalledTimes(2);
        expect(seen).toHaveLength(1);
    });

    it('reports declarations when the last procedure is deleted with the caret still at zero', () => {
        const one = editor(0, 'Sub Post()\nEnd Sub');
        setActiveEditor(one);
        track();
        change(one, '', 0, '', 1, false);
        expect(tracker?.current?.label).toBe('(Declarations)');
        expect(seen.map(position => position?.label)).toEqual(['(Declarations)']);
    });

    it('updates procedure ownership when a lead-in changes without a selection event', () => {
        const one = editor(6);
        setActiveEditor(one);
        track();
        change(one, SOURCE.replace('End Sub\n\nFunction', 'End Sub\nx = 1\nFunction'), 6, 'x = 1', 6, false);
        expect(tracker?.current?.label).toBe('Sub Post');
    });

    it('updates current range spans even when the procedure label stays the same', () => {
        const one = editor(3);
        setActiveEditor(one);
        track();
        const previous = tracker?.current?.procedure;
        change(one, SOURCE.replace('End Sub\n\nFunction', 'End Sub\nx = 1\nFunction'), 6, 'x = 1', 6, false);
        expect(tracker?.current?.label).toBe('Sub Post');
        expect(tracker?.current?.procedure).not.toBe(previous);
        expect(tracker?.current?.procedure?.lastLine).toBe(6);
        expect(seen).toEqual([]);
        expect(one.document.getText).toHaveBeenCalledTimes(1);
    });

    it('restores lead-in ownership after Backspace without rescanning source', () => {
        const one = editor(6);
        setActiveEditor(one);
        track();
        change(one, SOURCE.replace('End Sub\n\nFunction', 'End Sub\nx\nFunction'), 6, 'x', 6, false);
        expect(tracker?.current?.label).toBe('Sub Post');
        change(one, SOURCE, 6, '', 6, false);
        expect(tracker?.current?.label).toBe('Function Total');
        expect(seen.map(position => position?.label)).toEqual(['Sub Post', 'Function Total']);
        expect(one.document.getText).toHaveBeenCalledTimes(1);
    });

    it('keeps ordinary body edits cached and quiet without selection events', () => {
        const one = editor(3);
        setActiveEditor(one);
        track();
        change(one, SOURCE.replace('Debug.Print 1', 'Debug.Print 10'), 3, '0', 3, false);
        expect(one.document.getText).toHaveBeenCalledTimes(1);
        expect(seen).toEqual([]);
    });

    it('defers an inactive document rescan until it is shown again', () => {
        const one = editor(3), other = editor(8);
        setActiveEditor(one);
        track();
        switchTo(other);
        seen = [];
        change(one, SOURCE.replace('Sub Post()', 'Sub Posted()'), 2, 'ed', 2, false);
        expect(one.document.getText).toHaveBeenCalledTimes(1);
        expect(seen).toEqual([]);
        switchTo(one);
        expect(tracker?.current?.label).toBe('Sub Posted');
        expect(one.document.getText).toHaveBeenCalledTimes(2);
    });

    it('does not reread the module while editing body code', () => {
        const one = editor(3);
        setActiveEditor(one);
        track();
        change(one, SOURCE.replace('Debug.Print 1', 'Debug.Print 10'), 3, '0');
        change(one, SOURCE.replace('Debug.Print 1', 'Debug.Print 100'), 3, '0');
        expect(one.document.getText).toHaveBeenCalledTimes(1);
        expect(tracker?.current?.label).toBe('Sub Post');
        expect(seen).toEqual([]);
    });

    it('rescans a renamed header', () => {
        const one = editor(3);
        setActiveEditor(one);
        track();
        change(one, SOURCE.replace('Sub Post()', 'Sub Posted()'), 2, 'ed');
        expect(one.document.getText).toHaveBeenCalledTimes(2);
        expect(tracker?.current?.label).toBe('Sub Posted');
    });

    it('updates the next procedure lead-in boundary without a full source read', () => {
        const one = editor(6);
        setActiveEditor(one);
        track();
        expect(tracker?.current?.label).toBe('Function Total');
        change(one, SOURCE.replace('End Sub\n\nFunction', 'End Sub\nx = 1\nFunction'), 6, 'x = 1');
        expect(one.document.getText).toHaveBeenCalledTimes(1);
        expect(tracker?.current?.label).toBe('Sub Post');
    });

    it('rescans line insertion even when it is ordinary body code', () => {
        const one = editor(3);
        setActiveEditor(one);
        track();
        change(one, SOURCE.replace('Debug.Print 1', 'Debug.Print 1\n    x = 2'), 3, '\nx = 2');
        expect(one.document.getText).toHaveBeenCalledTimes(2);
    });
    it('keeps comment lead-in edits local', () => {
        const initial = SOURCE.replace('End Sub\n\nFunction', "End Sub\n' old\nFunction");
        const one = editor(6, initial);
        setActiveEditor(one);
        track();
        change(one, initial.replace("' old", "' updated"), 6, 'updated');
        expect(one.document.getText).toHaveBeenCalledTimes(1);
        expect(tracker?.current?.label).toBe('Function Total');
    });

    it('rescans when a document version was missed', () => {
        const one = editor(3);
        setActiveEditor(one);
        track();
        one.document.version++;
        change(one, SOURCE.replace('Sub Post()', 'Sub Posted()'), 3, '0');
        expect(one.document.getText).toHaveBeenCalledTimes(2);
        expect(tracker?.current?.label).toBe('Sub Posted');
    });

});
