import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
const host = vi.hoisted(() => ({ selection: undefined as unknown, reads: 0 }));
vi.mock('vscode', async () => {
    const base = await import('./helpers/vscodeMock');
    host.selection = new base.EventEmitter<unknown>();
    return base.vscodeMock({ window: {
        activeTextEditor: undefined,
        onDidChangeActiveTextEditor: new base.EventEmitter<unknown>().event,
        onDidChangeTextEditorSelection: (host.selection as { event: unknown }).event,
    } });
});
vi.mock('../src/vbaDocumentLocation', () => ({ moduleLocationOfDocument: () => ({ projectPath: 'Book.xlsm', moduleName: 'M', native: false }) }));
vi.mock('../src/vbaProcedureAtLine', async (importOriginal) => {
    const real = await importOriginal<typeof import('../src/vbaProcedureAtLine')>();
    return { ...real, vbaProcedureRanges: (source: string) => real.vbaProcedureRanges(source).map(range => ({
        ...range,
        get firstLine() { host.reads++; return range.firstLine; },
        get lastLine() { host.reads++; return range.lastLine; },
    })) };
});
import * as vscode from 'vscode';
import { VbaCaretProcedureTracker } from '../src/vbaCaretProcedure';
import { vbaProcedureAtLine, vbaProcedureLabel, vbaProcedureRanges } from '../src/vbaProcedureAtLine';
let tracker: VbaCaretProcedureTracker | undefined;
beforeEach(() => { host.reads = 0; });
afterEach(() => { tracker?.dispose(); tracker = undefined; });
function activate(source: string, line: number) {
    const document = { uri: { toString: () => 'xlide-vba:/M.bas' }, version: 1, getText: vi.fn(() => source) };
    const editor = { document, selection: { active: { line } } };
    (vscode.window as unknown as { activeTextEditor: unknown }).activeTextEditor = editor;
    tracker = new VbaCaretProcedureTracker();
    return editor;
}
function move(editor: ReturnType<typeof activate>, line: number) {
    editor.selection.active.line = line;
    (host.selection as { fire: (event: unknown) => void }).fire({ textEditor: editor });
}

describe('caret selection lookup', () => {
    it('bounds lookup work during 200 moves within the last of 1000 procedures', () => {
        const source = ['Option Explicit', ...Array.from({ length: 1000 }, (_, i) => 'Sub P' + i + '()\nDebug.Print 1\nEnd Sub')].join('\n');
        const editor = activate(source, 2998);
        host.reads = 0;
        for (let i = 0; i < 200; i++) {
            move(editor, 2998 + i % 3);
            expect(tracker?.current?.label).toBe('Sub P999');
        }
        expect(host.reads).toBeLessThanOrEqual(6000);
        expect(editor.document.getText).toHaveBeenCalledTimes(1);
    });

    it('matches the general helper across generated sources, boundaries and invalid lines', () => {
        for (let sample = 0; sample < 120; sample++) {
            const count = sample % 31, eol = ['\n', '\r\n', '\r'][sample % 3];
            const lines = ['Option Explicit', 'Dim declaration As Long'];
            for (let index = 0; index < count; index++) {
                const kind = ['Sub', 'Function', 'Property Get'][index % 3];
                lines.push('', index % 2 ? "' leading" : 'Rem leading', 'Private ' + kind + ' P' + index + '()', 'Debug.Print 1', 'End ' + kind.split(' ')[0]);
                if (index % 3 === 0) lines.push('unexpectedCode');
            }
            const source = lines.join(eol), ranges = vbaProcedureRanges(source);
            for (let index = 0; index < ranges.length; index++) {
                expect(ranges[index].lastLine).toBeGreaterThanOrEqual(ranges[index].firstLine);
                if (index) expect(ranges[index].firstLine).toBeGreaterThan(ranges[index - 1].lastLine);
            }
            const editor = activate(source, 0);
            const queries = [...Array.from({ length: lines.length * 2 + 8 }, (_, index) => index / 2 - 2), NaN, Infinity, -Infinity];
            for (const line of queries) {
                move(editor, line);
                expect(tracker?.current?.label, `sample ${sample}, line ${line}`).toBe(vbaProcedureLabel(vbaProcedureAtLine(ranges, line)));
            }
            expect(editor.document.getText).toHaveBeenCalledTimes(1);
            tracker?.dispose(); tracker = undefined;
        }
    });

    it('rebuilds lookup ranges after an edit changes procedure locations and names', () => {
        const editor = activate('Option Explicit\nSub Old()\nEnd Sub', 2);
        expect(tracker?.current?.label).toBe('Sub Old');
        editor.document.getText.mockReturnValue('Option Explicit\nDim x As Long\n\nSub New()\nEnd Sub');
        editor.document.version++;
        move(editor, 1); expect(tracker?.current?.label).toBe('(Declarations)');
        move(editor, 4); expect(tracker?.current?.label).toBe('Sub New');
        expect(editor.document.getText).toHaveBeenCalledTimes(2);
    });

    it('keeps first-match semantics for general unsorted or overlapping range inputs', () => {
        const ranges = [{ kind: 'Sub', name: 'First', firstLine: 10, lastLine: 20 }, { kind: 'Sub', name: 'Second', firstLine: 0, lastLine: 30 }];
        expect(vbaProcedureAtLine(ranges, 12)).toBe(ranges[0]);
        expect(vbaProcedureAtLine(ranges, 2)).toBe(ranges[1]);
    });
});
