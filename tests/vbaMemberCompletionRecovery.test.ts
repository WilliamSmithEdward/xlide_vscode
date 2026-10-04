import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type * as vscodeTypes from 'vscode';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
    window: { activeTextEditor: undefined, onDidChangeTextEditorSelection: vi.fn(() => ({ dispose: vi.fn() })) },
    workspace: { onDidCloseTextDocument: vi.fn(() => ({ dispose() {} })) },
}));

import * as vscode from 'vscode';
import { hasMemberCompletions, resolveMemberCompletions } from '../src/analyzer';
import { VbaMemberCompletionProvider } from '../src/vbaCompletionProvider';
import type { EditorProjectContext, VbaEditorProjectContextService } from '../src/vbaEditorProjectContext';

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
        localEditorProjectContext: vi.fn((): EditorProjectContext => ({})),
        cheapEditorProjectContext: vi.fn((): EditorProjectContext => ({})),
        buildEditorProjectContextWithin: vi.fn(async (): Promise<EditorProjectContext | undefined> => ({})),
        buildEditorProjectContext: vi.fn(async (): Promise<EditorProjectContext> => ({})),
    };
    const provider = new VbaMemberCompletionProvider(projectContext as unknown as VbaEditorProjectContextService);
    return { provider, document, editor, event, projectContext,
        send: () => provider.handleTextDocumentChange(event as unknown as vscodeTypes.TextDocumentChangeEvent) };
}

beforeEach(() => { vi.useFakeTimers(); vi.mocked(vscode.commands.executeCommand).mockClear(); });
afterEach(() => { vi.clearAllTimers(); vi.useRealTimers(); });

describe('member completion recovery on Backspace', () => {
    it.each(['ThisWorkbook.Sheets(1).[Ce', 'ThisWorkbook.Sheets(1).[Ce]'])(
        'recovers within a bracketed member: %s', async line => {
            const edit = deletion(line, { column: line.endsWith(']') ? line.length - 1 : line.length });
            edit.send();
            await vi.runAllTimersAsync();
            expect(vscode.commands.executeCommand).toHaveBeenCalledOnce();
        });
    it('waits for a delayed caret update and disposes its selection listener', async () => {
        const edit = deletion('ThisWorkbook.Sheets(1).ce');
        const expected = edit.editor.selection.active;
        edit.editor.selection.active = new vscode.Position(expected.line, expected.character + 1);
        const dispose = vi.fn();
        vi.mocked(vscode.window.onDidChangeTextEditorSelection).mockReturnValueOnce({ dispose });
        edit.send();
        await vi.advanceTimersByTimeAsync(0);
        expect(vscode.commands.executeCommand).not.toHaveBeenCalled();
        edit.editor.selection.active = expected;
        const listener = vi.mocked(vscode.window.onDidChangeTextEditorSelection).mock.calls.at(-1)![0];
        listener({ textEditor: edit.editor } as unknown as vscodeTypes.TextEditorSelectionChangeEvent);
        await vi.runAllTimersAsync();
        expect(vscode.commands.executeCommand).toHaveBeenCalledOnce();
        expect(dispose).toHaveBeenCalledOnce();
    });
    it('drops delayed recovery when the next selection moves elsewhere', async () => {
        const edit = deletion('ThisWorkbook.Sheets(1).ce');
        edit.editor.selection.active = new vscode.Position(edit.editor.selection.active.line, edit.editor.selection.active.character + 1);
        edit.send();
        await vi.advanceTimersByTimeAsync(0);
        edit.editor.selection.active = new vscode.Position(1, 0);
        const listener = vi.mocked(vscode.window.onDidChangeTextEditorSelection).mock.calls.at(-1)![0];
        listener({ textEditor: edit.editor } as unknown as vscodeTypes.TextEditorSelectionChangeEvent);
        edit.editor.selection.active = edit.event.contentChanges[0].range.start;
        listener({ textEditor: edit.editor } as unknown as vscodeTypes.TextEditorSelectionChangeEvent);
        await vi.runAllTimersAsync();
        expect(vscode.commands.executeCommand).not.toHaveBeenCalled();
    });
    it('expires and disposes recovery if the expected caret update never arrives', async () => {
        const edit = deletion('ThisWorkbook.Sheets(1).ce');
        const expected = edit.editor.selection.active;
        edit.editor.selection.active = new vscode.Position(expected.line, expected.character + 1);
        const dispose = vi.fn();
        vi.mocked(vscode.window.onDidChangeTextEditorSelection).mockReturnValueOnce({ dispose });
        edit.send();
        await vi.runAllTimersAsync();
        expect(dispose).toHaveBeenCalledOnce();
        edit.editor.selection.active = expected;
        const listener = vi.mocked(vscode.window.onDidChangeTextEditorSelection).mock.calls.at(-1)![0];
        listener({ textEditor: edit.editor } as unknown as vscodeTypes.TextEditorSelectionChangeEvent);
        expect(vscode.commands.executeCommand).not.toHaveBeenCalled();
    });
    it('loads source members that extend a known Me host surface', async () => {
        const edit = deletion('Me.He');
        edit.projectContext.cheapEditorProjectContext.mockReturnValue({ meType: 'Excel.Workbook', meProjectType: 'ThisWorkbook' });
        edit.projectContext.buildEditorProjectContext.mockResolvedValue({ meType: 'Excel.Workbook', meProjectType: 'ThisWorkbook',
            projectClassMembers: [{ name: 'ThisWorkbook', kind: 'document', moduleName: 'ThisWorkbook',
                members: [{ name: 'Hello', kind: 'method', moduleName: 'ThisWorkbook' }] }] });
        edit.send();
        await vi.runAllTimersAsync();
        expect(vscode.commands.executeCommand).toHaveBeenCalledOnce();
    });

    it('does not load project context for a known host receiver with no matching member', async () => {
        const edit = deletion('ThisWorkbook.Sheets(1).cez');
        edit.send();
        await vi.runAllTimersAsync();
        expect(vscode.commands.executeCommand).not.toHaveBeenCalled();
        expect(edit.projectContext.buildEditorProjectContextWithin).not.toHaveBeenCalled();
        expect(edit.projectContext.buildEditorProjectContext).not.toHaveBeenCalled();
    });
    it('recovers project members when loading takes longer than the completion budget', async () => {
        const edit = deletion('obj.He', { prelude: 'Sub Demo()\nDim obj As Widget\n' });
        const context: EditorProjectContext = { projectClassMembers: [{ name: 'Widget', kind: 'class', moduleName: 'Widget',
            members: [{ name: 'Hello', kind: 'method', moduleName: 'Widget' }] }] };
        edit.projectContext.buildEditorProjectContextWithin.mockImplementation(() =>
            new Promise(resolve => setTimeout(() => resolve(undefined), 150)));
        edit.projectContext.buildEditorProjectContext.mockImplementation(() =>
            new Promise(resolve => setTimeout(() => resolve(context), 250)));
        edit.send();
        await vi.runAllTimersAsync();
        expect(vscode.commands.executeCommand).toHaveBeenCalledOnce();
    });

    it.each(['version', 'move', 'switch'])('drops recovery superseded during project loading: %s', async action => {
        const edit = deletion('obj.He', { prelude: 'Sub Demo()\nDim obj As Widget\n' });
        edit.projectContext.buildEditorProjectContext.mockImplementation(async () => {
            if (action === 'version') { edit.document.version++; }
            if (action === 'move') { edit.editor.selection.active = new vscode.Position(2, 0); }
            if (action === 'switch') { (vscode.window as unknown as { activeTextEditor: unknown }).activeTextEditor = undefined; }
            return { projectClassMembers: [{ name: 'Widget', kind: 'class', moduleName: 'Widget',
                members: [{ name: 'Hello', kind: 'method', moduleName: 'Widget' }] }] };
        });
        edit.send();
        await vi.runAllTimersAsync();
        expect(vscode.commands.executeCommand).not.toHaveBeenCalled();
    });
    it('recovers Unicode members ending in a combining mark', async () => {
        const edit = deletion('obj.का', { prelude: 'Sub Demo()\nDim obj As Widget\n' });
        edit.projectContext.buildEditorProjectContext.mockResolvedValue({
            projectClassMembers: [{ name: 'Widget', kind: 'class', moduleName: 'Widget',
                members: [{ name: 'काम', kind: 'property', moduleName: 'Widget' }] }],
        });
        edit.send();
        await vi.runAllTimersAsync();
        expect(vscode.commands.executeCommand).toHaveBeenCalledOnce();
    });
    it('probes matching members without rendering completion documentation', () => {
        const source = 'Sub Demo()\nMe.He\nEnd Sub';
        const context = { meProjectType: 'Widget', projectClassMembers: [{ name: 'Widget', kind: 'class' as const,
            moduleName: 'Widget', members: [{ name: 'Hello', kind: 'method' as const, moduleName: 'Widget',
                get doc(): never { throw new Error('documentation should not be read'); } }] }] };
        expect(hasMemberCompletions(source, source.indexOf('Me.He') + 5, context)).toBe(true);
    });

    it('recovers cross-module members when the full context is initially cold', async () => {
        const edit = deletion('obj.He', { prelude: 'Sub Demo()\nDim obj As Widget\n' });
        edit.projectContext.buildEditorProjectContext.mockResolvedValue({
            projectClassMembers: [{ name: 'Widget', kind: 'class', moduleName: 'Widget',
                members: [{ name: 'Hello', kind: 'method', moduleName: 'Widget' }] }],
        });
        edit.send();
        await vi.runAllTimersAsync();
        expect(vscode.commands.executeCommand).toHaveBeenCalledWith('editor.action.triggerSuggest', { auto: true });
    });
    it('does not build full module/project contexts just to reopen host members', async () => {
        const edit = deletion('ThisWorkbook.Sheets(1).ce');
        edit.send();
        await vi.runAllTimersAsync();
        expect(vscode.commands.executeCommand).toHaveBeenCalledOnce();
        expect(edit.projectContext.localEditorProjectContext).not.toHaveBeenCalled();
        expect(edit.projectContext.buildEditorProjectContextWithin).not.toHaveBeenCalled();
    });

    it('reopens Cells after ThisWorkbook.Sheets(1).cez is shortened to .ce', async () => {
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
        await vi.runAllTimersAsync();
        expect(vscode.commands.executeCommand).toHaveBeenCalledWith('editor.action.triggerSuggest', { auto: true });
    });

    it.each([
        ['ThisWorkbook.Sheets(1).', undefined],
        ['    .ce', 'Sub Demo()\nWith ThisWorkbook.Sheets(1)\n'],
        ['    sheet.ce', 'Sub Demo()\nDim sheet As Worksheet\n'],
    ])('reopens members for %s', async (line, prelude) => {
        deletion(line, { prelude }).send();
        await vi.runAllTimersAsync();
        expect(vscode.commands.executeCommand).toHaveBeenCalledOnce();
    });

    it.each([
        'ThisWorkbook.Sheets(1).cez',
        "'ThisWorkbook.Sheets(1).ce",
        'Debug.Print "ThisWorkbook.Sheets(1).ce',
        'unknown.ce',
        'ce',
        'ThisWorkbook.Sheets(1)',
    ])('does not open suggestions for %s', async line => {
        deletion(line).send();
        await vi.runAllTimersAsync();
        expect(vscode.commands.executeCommand).not.toHaveBeenCalled();
    });

    it('ignores undo, multi-character deletion, replacement and multiple cursors', async () => {
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
        await vi.runAllTimersAsync();
        expect(vscode.commands.executeCommand).not.toHaveBeenCalled();
    });

    it.each(['move', 'switch', 'version', 'close', 'select', 'delete'])('drops pending recovery after %s', async action => {
        const edit = deletion('ThisWorkbook.Sheets(1).ce');
        edit.send();
        if (action === 'move') { edit.editor.selection.active = new vscode.Position(1, 0); }
        if (action === 'switch') { (vscode.window as unknown as { activeTextEditor: unknown }).activeTextEditor = undefined; }
        if (action === 'version') { edit.document.version++; }
        if (action === 'close') { edit.document.isClosed = true; }
        if (action === 'select') { edit.editor.selection.isEmpty = false; }
        if (action === 'delete') { edit.editor.selection.active = new vscode.Position(1, 24); }
        await vi.runAllTimersAsync();
        expect(vscode.commands.executeCommand).not.toHaveBeenCalled();
        expect(edit.projectContext.localEditorProjectContext).not.toHaveBeenCalled();
    });
});
