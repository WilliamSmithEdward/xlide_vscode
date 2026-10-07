import { afterEach, describe, expect, it, vi } from 'vitest';
import { writeFileSync } from 'node:fs';
import { lexerStrippedLines } from '../src/analyzer/lexer/strippedLines';
import { tokenizeCached } from '../src/analyzer/lexer/tokenize';

const selectionListeners = vi.hoisted(() => new Set<(event: import('vscode').TextEditorSelectionChangeEvent) => unknown>());

vi.mock('vscode', async () => ({
    ...(await import('./helpers/vscodeMock')).vscodeMock({ window: {
        onDidChangeTextEditorSelection: vi.fn(listener => {
            selectionListeners.add(listener);
            return { dispose: () => selectionListeners.delete(listener) };
        }),
    } }),
    EndOfLine: { LF: 1, CRLF: 2 },
    Selection: class { constructor(public anchor: unknown, public active: unknown) {} },
}));
import * as vscode from 'vscode';
import { registerVbaAutoBlock, registerVbaLoopIteratorSync } from '../src/vbaTypingAutomation';

afterEach(async () => {
    if (vi.isFakeTimers()) { await vi.runOnlyPendingTimersAsync(); }
    else { await new Promise(resolve => setTimeout(resolve, 0)); }
    vi.useRealTimers(); selectionListeners.clear(); vi.restoreAllMocks();
});

function enterScenario(source: string, previousLine: number) {
    const lines = source.split('\n');
    const document = {
        uri: vscode.Uri.file('/enter.bas'), languageId: 'vba', version: 1, isClosed: false, eol: vscode.EndOfLine.LF,
        lineCount: lines.length, lineAt: (line: number) => ({ text: lines[line] }),
        getText: vi.fn(() => lines.join('\n')),
        offsetAt: (position: vscode.Position) => lines.slice(0, position.line).reduce((total, line) => total + line.length + 1, 0) + position.character,
        positionAt: (offset: number) => {
            let line = 0;
            while (line + 1 < lines.length && offset > lines[line].length) { offset -= lines[line++].length + 1; }
            return new vscode.Position(line, offset);
        },
    } as unknown as vscode.TextDocument;
    const replacements: string[] = [];
    const editor = {
        document,
        selection: { isEmpty: true, anchor: new vscode.Position(previousLine + 1, 0), active: new vscode.Position(previousLine + 1, 0) },
        edit: vi.fn(async (apply: (edit: unknown) => void) => {
            apply({ replace: (_range: unknown, text: string) => replacements.push(text), insert: vi.fn() });
            Object.assign(document, { version: document.version + 1 });
            return true;
        }),
    };
    Object.assign(vscode.window, { activeTextEditor: editor });
    const context = { subscriptions: [] } as unknown as vscode.ExtensionContext;
    registerVbaAutoBlock(context);
    const listener = vi.mocked(vscode.workspace.onDidChangeTextDocument).mock.calls.at(-1)![0];
    const event = {
        document, contentChanges: [{
            range: new vscode.Range(new vscode.Position(previousLine, lines[previousLine].length),
                new vscode.Position(previousLine, lines[previousLine].length)),
            text: '\n',
        }],
    } as unknown as vscode.TextDocumentChangeEvent;
    return { document, editor, replacements, lines, event, listener, invoke: async () => { await listener(event); },
        dispose: () => context.subscriptions.forEach(item => item.dispose()) };
}

describe('Smart Enter surface work', () => {
    it('continues processing a different module while an earlier module edit is pending', async () => {
        const first = enterScenario('Sub A()\n    If ready Then\n\nEnd Sub\n', 1);
        const second = enterScenario('Sub B()\n    While ready\n\nEnd Sub\n', 1);
        let finish!: (value: boolean) => void;
        first.editor.edit.mockImplementationOnce(() => new Promise<boolean>(resolve => { finish = resolve; }));
        Object.assign(vscode.window, { activeTextEditor: first.editor });
        const pending = first.invoke();
        try {
            await first.invoke();
            expect(first.editor.edit).toHaveBeenCalledTimes(1);
            Object.assign(vscode.window, { activeTextEditor: second.editor });
            await first.listener(second.event);
            expect(second.editor.edit).toHaveBeenCalledTimes(1);
            expect(second.replacements[0]).toContain('Wend');
        } finally { finish(true); await pending; first.dispose(); second.dispose(); }
    });

    it('synchronizes another module while an earlier loop rename edit is pending', async () => {
        const first = enterScenario('Sub A()\n    For i = 1 To 10\n    Next j\nEnd Sub\n', 1);
        const second = enterScenario('Sub B()\n    For k = 1 To 10\n    Next j\nEnd Sub\n', 1);
        const context = { subscriptions: [] } as unknown as vscode.ExtensionContext;
        registerVbaLoopIteratorSync(context);
        const listener = vi.mocked(vscode.workspace.onDidChangeTextDocument).mock.calls.at(-1)![0];
        const eventFor = (document: vscode.TextDocument, letter: string) => ({
            document, contentChanges: [{ text: letter, range: new vscode.Range(new vscode.Position(1, 8), new vscode.Position(1, 9)) }],
        }) as unknown as vscode.TextDocumentChangeEvent;
        let finish!: (value: boolean) => void;
        first.editor.edit.mockImplementationOnce(() => new Promise<boolean>(resolve => { finish = resolve; }));
        Object.assign(vscode.window, { activeTextEditor: first.editor });
        const pending = listener(eventFor(first.document, 'i'));
        try {
            expect(first.editor.edit).toHaveBeenCalledTimes(1);
            await listener(eventFor(first.document, 'i'));
            expect(first.editor.edit).toHaveBeenCalledTimes(1);
            Object.assign(vscode.window, { activeTextEditor: second.editor });
            await listener(eventFor(second.document, 'k'));
            expect(second.editor.edit).toHaveBeenCalledTimes(1);
            expect(second.replacements).toEqual(['k']);
        } finally { finish(true); await pending; context.subscriptions.forEach(item => item.dispose()); first.dispose(); second.dispose(); }
    });

    it('does not scan an inactive module after a block-opener newline', async () => {
        const scenario = enterScenario('Sub T()\n    If ready Then\n\nEnd Sub\n', 1);
        try {
            Object.assign(vscode.window, { activeTextEditor: undefined });
            await scenario.invoke();
            expect(scenario.document.getText).not.toHaveBeenCalled();
            expect(scenario.editor.edit).not.toHaveBeenCalled();
        } finally { scenario.dispose(); }
    });

    it.each([1, 2])('does not generate edits for undo/redo reason %s', async reason => {
        const scenario = enterScenario('Sub T()\n    If ready Then\n\nEnd Sub\n', 1);
        try {
            Object.assign(scenario.event, { reason });
            await scenario.invoke();
            expect(scenario.document.getText).not.toHaveBeenCalled();
            expect(scenario.editor.edit).not.toHaveBeenCalled();
        } finally { scenario.dispose(); }
    });

    it('leaves the caret alone when the smart-block edit is rejected', async () => {
        const scenario = enterScenario('Sub T()\n    If ready Then\n    \t\n    End If\nEnd Sub\n', 1);
        try {
            const initial = scenario.editor.selection;
            scenario.editor.edit.mockResolvedValueOnce(false);
            await scenario.invoke();
            expect(scenario.editor.selection).toBe(initial);
        } finally { scenario.dispose(); }
    });

    it('does not move the caret back after the user navigates away before the delayed pass', async () => {
        vi.useFakeTimers();
        const scenario = enterScenario('Sub T()\n    With ActiveSheet\n    \t\n    End With\nEnd Sub\n', 1);
        try {
            scenario.editor.edit.mockImplementationOnce(async () => {
                scenario.lines[2] = '    \t.';
                Object.assign(scenario.document, { version: scenario.document.version + 1 });
                return true;
            });
            await scenario.invoke();
            expect(scenario.editor.selection.active.line).toBe(2);
            scenario.editor.selection = { isEmpty: true, anchor: new vscode.Position(0, 0), active: new vscode.Position(0, 0) };
            await vi.advanceTimersByTimeAsync(0);
            expect(scenario.editor.selection.active.line).toBe(0);
        } finally { scenario.dispose(); }
    });

    it('does not open dot suggestions after the caret moved left from the seeded dot', async () => {
        vi.useFakeTimers();
        const scenario = enterScenario('Sub T()\n    With ActiveSheet\n    \t\n    End With\nEnd Sub\n', 1);
        try {
            scenario.editor.edit.mockImplementationOnce(async () => {
                scenario.lines[2] = '    \t.';
                Object.assign(scenario.document, { version: scenario.document.version + 1 });
                return true;
            });
            await scenario.invoke();
            scenario.editor.selection = {
                isEmpty: true, anchor: new vscode.Position(2, 5), active: new vscode.Position(2, 5),
            };
            await vi.advanceTimersByTimeAsync(0);
            expect(vscode.commands.executeCommand).not.toHaveBeenCalledWith('editor.action.triggerSuggest');
        } finally { scenario.dispose(); }
    });

    it.each(['    value = 1', '    Debug.Print "hello"', "    value = 1 ' trailing", '    Rem comment'])(
        'does not read the module for Enter after %s', async line => {
            const scenario = enterScenario('Sub T()\n' + line + '\n\nEnd Sub\n', 1);
            try {
                await scenario.invoke();
                expect(scenario.document.getText).not.toHaveBeenCalled();
                expect(scenario.editor.edit).not.toHaveBeenCalled();
            } finally { scenario.dispose(); }
        });

    it('continues whole-line comments without reading the module', async () => {
        const scenario = enterScenario("Sub T()\n    '''  note\n\nEnd Sub\n", 1);
        try {
            await scenario.invoke();
            expect(scenario.document.getText).not.toHaveBeenCalled();
            expect(scenario.replacements).toEqual(["    '''  "]);
        } finally { scenario.dispose(); }
    });

    it('still checks the module for an active With before inserting a dot', async () => {
        const scenario = enterScenario('Sub T()\n    With ActiveSheet\n        .Name = "test"\n\n    End With\nEnd Sub\n', 2);
        try {
            await scenario.invoke();
            expect(scenario.document.getText).toHaveBeenCalledTimes(1);
            expect(scenario.replacements).toEqual(['        .']);
        } finally { scenario.dispose(); }
    });

    it('does not seed a dot outside a With', async () => {
        const scenario = enterScenario('Sub T()\n    .Name = "test"\n\nEnd Sub\n', 1);
        try {
            await scenario.invoke();
            expect(scenario.replacements).toEqual([]);
        } finally { scenario.dispose(); }
    });

    it('still inserts a closer after a new block opener', async () => {
        const scenario = enterScenario('Sub T()\n    If ready Then\n\nEnd Sub\n', 1);
        try {
            await scenario.invoke();
            expect(scenario.document.getText).toHaveBeenCalledTimes(1);
            expect(scenario.replacements[0]).toContain('End If');
        } finally { scenario.dispose(); }
    });

    it.each(['\n', '\r\n', '\r'])('preserves UTF-16 columns and continued comments using %j', eol => {
        const lines = ['Debug.Print "😀"; "ไทย": Rem hidden _', 'End If', 'x = #1/1/2026#'];
        const source = lines.join(eol);
        // Use token spans as an independent mask; UTF-16 offsets require code units.
        const units = source.split('');
        for (const token of tokenizeCached(source)) {
            if (token.kind !== 'comment' && token.kind !== 'stringLiteral') { continue; }
            for (let offset = token.start; offset < token.end; offset++) {
                if (units[offset] !== '\r' && units[offset] !== '\n') { units[offset] = ' '; }
            }
        }
        expect(lexerStrippedLines(source)).toEqual(units.join('').split(/\r\n|\r|\n/));
        expect(lexerStrippedLines(source).map(line => line.length)).toEqual(lines.map(line => line.length));
    });

    it.skipIf(!process.env.XLIDE_ENTER_BENCHMARK_OUTPUT)('measures Enter and stripped-line work', async () => {
        const source = Array.from({ length: 1200 }, (_, index) =>
            'Sub Probe' + index + '()\n    Debug.Print "probe"; "' + index + '" \' comment\nEnd Sub\n').join('\n');
        const plain = enterScenario(source + '\nSub Active()\n    value = 1\n\nEnd Sub\n', source.split('\n').length + 1);
        const comment = enterScenario(source + "\nSub Active()\n    ' note\n\nEnd Sub\n", source.split('\n').length + 1);
        const medians: Record<string, number> = {};
        const measure = async (name: string, action: () => unknown) => {
            await action();
            const times: number[] = [];
            for (let sample = 0; sample < 21; sample++) {
                const start = performance.now();
                await action();
                times.push(performance.now() - start);
            }
            medians[name] = times.sort((a, b) => a - b)[10];
        };
        try {
            await measure('strippedLinesMs', () => lexerStrippedLines(source));
            await measure('ordinaryEnterMs', () => plain.invoke());
            await measure('commentEnterMs', () => comment.invoke());
            writeFileSync(process.env.XLIDE_ENTER_BENCHMARK_OUTPUT!, JSON.stringify({ procedures: 1200, bytes: source.length, samples: 21, medians }, null, 2));
        } finally { plain.dispose(); comment.dispose(); }
    });
});


describe('pending Enter caret ownership', () => {
    for (const [kind, source, previousLine] of [
        ['block', 'Sub T()\n    With ActiveSheet\n    \t\n    End With\nEnd Sub\n', 1],
        ['comment', "Sub T()\n    ' note\n\nEnd Sub\n", 1],
        ['member', 'Sub T()\n    With ActiveSheet\n        .Name = "test"\n\n    End With\nEnd Sub\n', 2],
    ] as const) {
        for (const action of ['stay', 'up', 'keyboard', 'mouse', 'command', 'typing'] as const) {
            it(`retains ${action} ownership after a pending ${kind} edit`, async () => {
                vi.useFakeTimers();
                const scenario = enterScenario(source, previousLine);
                let finish!: (value: boolean) => void;
                scenario.editor.edit.mockImplementationOnce((apply: (edit: unknown) => void) => {
                    apply({ replace: (_range: unknown, text: string) => { scenario.lines[previousLine + 1] = text; }, insert: vi.fn() });
                    Object.assign(scenario.document, { version: scenario.document.version + 1 });
                    return new Promise<boolean>(resolve => { finish = resolve; });
                });
                try {
                    const pending = scenario.invoke();
                    for (let i = 0; i < 8 && !finish; i++) { await Promise.resolve(); }
                    expect(finish).toBeDefined();
                    if (action === 'up') {
                        scenario.editor.selection = { isEmpty: true, anchor: new vscode.Position(0, 0), active: new vscode.Position(0, 0) };
                    } else if (action === 'typing') {
                        Object.assign(scenario.document, { version: scenario.document.version + 1 });
                    } else if (action !== 'stay') {
                        // Even navigation ending at the original input position
                        // relinquishes ownership; position equality is insufficient.
                        for (const listener of selectionListeners) {
                            listener({ textEditor: scenario.editor, selections: [scenario.editor.selection],
                                kind: action === 'keyboard' ? 1 : action === 'mouse' ? 2 : 3,
                            } as unknown as vscode.TextEditorSelectionChangeEvent);
                        }
                    }
                    const selected = scenario.editor.selection;
                    finish(true); await pending; await vi.runOnlyPendingTimersAsync();
                    if (action === 'stay') {
                        expect(scenario.editor.selection.active.line).toBe(previousLine + 1);
                        expect(scenario.editor.selection.active.character).toBe(scenario.lines[previousLine + 1].length);
                    } else { expect(scenario.editor.selection).toBe(selected); }
                    expect(selectionListeners.size).toBe(0);
                } finally { scenario.dispose(); }
            });
        }
    }
    it.each(['reject', 'throw', 'closed', 'editor-switch'] as const)('disposes pending placement after %s', async action => {
        vi.useFakeTimers();
        const scenario = enterScenario('Sub T()\n    With ActiveSheet\n    \t\n    End With\nEnd Sub\n', 1);
        const selected = scenario.editor.selection;
        let finish!: (value: boolean) => void;
        let fail!: (reason: Error) => void;
        scenario.editor.edit.mockImplementationOnce(() => new Promise<boolean>((resolve, reject) => { finish = resolve; fail = reject; }));
        try {
            const pending = scenario.invoke();
            expect(selectionListeners.size).toBe(1);
            if (action === 'throw') {
                const assertion = expect(pending).rejects.toThrow('edit failure');
                fail(new Error('edit failure')); await assertion;
            } else {
                if (action === 'closed') { Object.assign(scenario.document, { isClosed: true, version: 2 }); }
                if (action === 'editor-switch') { Object.assign(vscode.window, { activeTextEditor: undefined }); Object.assign(scenario.document, { version: 2 }); }
                finish(action !== 'reject'); await pending; await vi.runOnlyPendingTimersAsync();
            }
            expect(scenario.editor.selection).toBe(selected);
            expect(selectionListeners.size).toBe(0);
        } finally { scenario.dispose(); }
    });
});
