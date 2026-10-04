import { afterEach, describe, expect, it, vi } from 'vitest';
import { writeFileSync } from 'node:fs';
import { lexerStrippedLines } from '../src/analyzer/lexer/strippedLines';
import { tokenizeCached } from '../src/analyzer/lexer/tokenize';

vi.mock('vscode', async () => ({
    ...(await import('./helpers/vscodeMock')).vscodeMock(),
    EndOfLine: { LF: 1, CRLF: 2 },
    Selection: class { constructor(public anchor: unknown, public active: unknown) {} },
}));
import * as vscode from 'vscode';
import { registerVbaAutoBlock } from '../src/vbaTypingAutomation';

afterEach(() => vi.restoreAllMocks());

function enterScenario(source: string, previousLine: number) {
    const lines = source.split('\n');
    const document = {
        uri: vscode.Uri.file('/enter.bas'), languageId: 'vba', eol: vscode.EndOfLine.LF,
        lineCount: lines.length, lineAt: (line: number) => ({ text: lines[line] }),
        getText: vi.fn(() => lines.join('\n')),
    } as unknown as vscode.TextDocument;
    const replacements: string[] = [];
    const editor = {
        document,
        selection: { isEmpty: true, active: new vscode.Position(previousLine + 1, 0) },
        edit: vi.fn(async (apply: (edit: unknown) => void) => {
            apply({ replace: (_range: unknown, text: string) => replacements.push(text), insert: vi.fn() });
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
    return { document, editor, replacements, invoke: async () => { await listener(event); },
        dispose: () => context.subscriptions.forEach(item => item.dispose()) };
}

describe('Smart Enter surface work', () => {
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
