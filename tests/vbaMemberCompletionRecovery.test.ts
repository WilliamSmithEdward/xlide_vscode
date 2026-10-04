import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type * as vscodeTypes from 'vscode';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
    window: { activeTextEditor: undefined },
    workspace: { onDidCloseTextDocument: vi.fn(() => ({ dispose() {} })) },
}));

import * as vscode from 'vscode';
import { resolveMemberCompletions } from '../src/analyzer';
import { VbaMemberCompletionProvider } from '../src/vbaCompletionProvider';
import type { VbaEditorProjectContextService } from '../src/vbaEditorProjectContext';

function deletion(line: string, options: { column?: number; prelude?: string } = {}) {
    const prelude = options.prelude ?? 'Sub Demo()\n';
    const source = prelude + line + '\nEnd Sub';
    const lineIndex = prelude.split('\n').length - 1;
    const column = options.column ?? line.length;
    const caret = new vscode.Position(lineIndex, column);
    const document = {
        uri: { scheme: 'file', path: '/Module1.bas' },
        languageId: 'vba', version: 2, isClosed: false,
        getText: () => source,
        lineAt: (index: number) => ({ text: source.split('\n')[index] }),
        offsetAt: (position: vscode.Position) =>
            source.split('\n').slice(0, position.line).reduce((total, text) => total + text.length + 1, 0) + position.character,
    };
    const selection = { active: caret, isEmpty: true };
    const editor = { document, selection, selections: [selection] };
    (vscode.window as unknown as { activeTextEditor: unknown }).activeTextEditor = editor;
    const event = {
        document,
        contentChanges: [{ text: '', rangeLength: 1,
            range: new vscode.Range(caret, new vscode.Position(lineIndex, column + 1)), rangeOffset: document.offsetAt(caret) }],
    };
    const projectContext = {
        cachedEditorProjectContext: vi.fn(() => undefined),
        localEditorProjectContext: vi.fn(() => ({})),
    };
    const provider = new VbaMemberCompletionProvider(projectContext as unknown as VbaEditorProjectContextService);
    return { provider, document, editor, event, projectContext,
        send: () => provider.handleTextDocumentChange(event as unknown as vscodeTypes.TextDocumentChangeEvent) };
}

beforeEach(() => { vi.useFakeTimers(); vi.mocked(vscode.commands.executeCommand).mockClear(); });
afterEach(() => { vi.clearAllTimers(); vi.useRealTimers(); });

describe('member completion recovery on Backspace', () => {
    it('reopens Cells after ThisWorkbook.Sheets(1).cez is shortened to .ce', () => {
        const before = 'Sub Demo()\nThisWorkbook.Sheets(1).cez\nEnd Sub';
        expect(resolveMemberCompletions(before, before.indexOf('cez') + 3)).toEqual([]);
        const edit = deletion('ThisWorkbook.Sheets(1).ce');
        expect(resolveMemberCompletions(edit.document.getText(), edit.document.offsetAt(edit.editor.selection.active))
            .map(member => member.name)).toContain('Cells');
        // VS Code updates the caret after its document-change notification.
        const caret = edit.editor.selection.active;
        edit.editor.selection.active = new vscode.Position(caret.line, caret.character + 1);
        edit.send();
        expect(vscode.commands.executeCommand).not.toHaveBeenCalled();
        edit.editor.selection.active = caret;
        vi.runAllTimers();
        expect(vscode.commands.executeCommand).toHaveBeenCalledWith('editor.action.triggerSuggest');
    });

    it.each([
        ['ThisWorkbook.Sheets(1).', undefined],
        ['    .ce', 'Sub Demo()\nWith ThisWorkbook.Sheets(1)\n'],
        ['    sheet.ce', 'Sub Demo()\nDim sheet As Worksheet\n'],
    ])('reopens members for %s', (line, prelude) => {
        deletion(line, { prelude }).send();
        vi.runAllTimers();
        expect(vscode.commands.executeCommand).toHaveBeenCalledOnce();
    });

    it.each([
        'ThisWorkbook.Sheets(1).cez',
        "'ThisWorkbook.Sheets(1).ce",
        'Debug.Print "ThisWorkbook.Sheets(1).ce',
        'unknown.ce',
        'ce',
        'ThisWorkbook.Sheets(1)',
    ])('does not open suggestions for %s', line => {
        deletion(line).send();
        vi.runAllTimers();
        expect(vscode.commands.executeCommand).not.toHaveBeenCalled();
    });

    it('ignores undo, multi-character deletion, replacement and multiple cursors', () => {
        const edit = deletion('ThisWorkbook.Sheets(1).ce');
        for (const patch of [
            { reason: 1 },
            { contentChanges: [{ ...edit.event.contentChanges[0], rangeLength: 2 }] },
            { contentChanges: [{ ...edit.event.contentChanges[0], text: 'a' }] },
            { contentChanges: [edit.event.contentChanges[0], edit.event.contentChanges[0]] },
        ]) {
            edit.provider.handleTextDocumentChange({ ...edit.event, ...patch } as unknown as vscodeTypes.TextDocumentChangeEvent);
        }
        edit.editor.selections.push(edit.editor.selection);
        edit.send();
        vi.runAllTimers();
        expect(vscode.commands.executeCommand).not.toHaveBeenCalled();
    });

    it.each(['move', 'switch', 'version', 'close', 'select', 'delete'])('drops pending recovery after %s', action => {
        const edit = deletion('ThisWorkbook.Sheets(1).ce');
        edit.send();
        if (action === 'move') { edit.editor.selection.active = new vscode.Position(1, 0); }
        if (action === 'switch') { (vscode.window as unknown as { activeTextEditor: unknown }).activeTextEditor = undefined; }
        if (action === 'version') { edit.document.version++; }
        if (action === 'close') { edit.document.isClosed = true; }
        if (action === 'select') { edit.editor.selection.isEmpty = false; }
        if (action === 'delete') { edit.editor.selection.active = new vscode.Position(1, 24); }
        vi.runAllTimers();
        expect(vscode.commands.executeCommand).not.toHaveBeenCalled();
        expect(edit.projectContext.localEditorProjectContext).not.toHaveBeenCalled();
    });
});
