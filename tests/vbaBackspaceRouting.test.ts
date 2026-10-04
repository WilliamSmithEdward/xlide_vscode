import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type * as vscodeTypes from 'vscode';
import { readFileSync } from 'node:fs';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({ window: {
    onDidChangeActiveTextEditor: vi.fn(() => ({ dispose() {} })),
    onDidChangeTextEditorSelection: vi.fn(() => ({ dispose() {} })),
} }));
import * as vscode from 'vscode';
import { backspaceNeedsExtension, BACKSPACE_NEEDS_EXTENSION_CONTEXT, registerVbaEditorCommands } from '../src/vbaEditorCommands';

function editorFor(lines: string[], line = lines.length - 1, character = lines[line].length) {
    const position = new vscode.Position(line, character);
    const selection = { active: position, anchor: position, isEmpty: true };
    return {
        document: { languageId: 'vba', uri: vscode.Uri.file('/backspace.bas'), isClosed: false, version: 1, lineCount: lines.length,
            lineAt: vi.fn((index: number) => ({ text: lines[index] })), getText: vi.fn(() => { throw new Error('Context must not read the whole module'); }) },
        selection, selections: [selection],
    } as unknown as vscodeTypes.TextEditor;
}
function activate(editor: vscodeTypes.TextEditor | undefined) {
    Object.assign(vscode.window, { activeTextEditor: editor });
    const context = { subscriptions: [] } as unknown as vscodeTypes.ExtensionContext;
    registerVbaEditorCommands(context);
    contexts.push(context);
}
const contexts: vscodeTypes.ExtensionContext[] = [];
function latestValue() { return vi.mocked(vscode.commands.executeCommand).mock.calls.at(-1); }
beforeEach(() => { vi.mocked(vscode.commands.registerCommand).mockClear(); vi.mocked(vscode.commands.executeCommand).mockClear(); vi.mocked(vscode.window.onDidChangeActiveTextEditor).mockClear(); vi.mocked(vscode.window.onDidChangeTextEditorSelection).mockClear(); vi.mocked(vscode.workspace.onDidChangeTextDocument).mockClear(); vi.mocked(vscode.workspace.onDidCloseTextDocument).mockClear(); });
afterEach(() => { contexts.splice(0).forEach(context => context.subscriptions.forEach(item => item.dispose())); vi.restoreAllMocks(); });

describe('native Backspace routing', () => {
    it.each(['ThisWorkbook.Sheets(1).a', 'ThisWorkbook.Sheets(1).cez', 'value = 1', "' actual comment", ''])('keeps ordinary deletion native on %j', line => {
        const editor = editorFor(['Sub Demo()', line]);
        expect(backspaceNeedsExtension(editor)).toBe(false);
        expect(editor.document.getText).not.toHaveBeenCalled();
    });
    it.each(['    ', '\t\t', ' \t '])('keeps whole-indent cleanup on %j', line => {
        expect(backspaceNeedsExtension(editorFor([line]))).toBe(true);
        expect(backspaceNeedsExtension(editorFor([line], 0, 0))).toBe(false);
    });
    it.each(["' ", "'' ", "''' ", "''''"])('keeps continued-comment cleanup for %j', line => {
        expect(backspaceNeedsExtension(editorFor(["'''' preceding", line]))).toBe(true);
        expect(backspaceNeedsExtension(editorFor(['value = 1', line]))).toBe(false);
        expect(backspaceNeedsExtension(editorFor([line]))).toBe(false);
    });
    it.each(['selection', 'multi', 'closed', 'language', 'staleCaret'])('leaves %s deletion to the editor', kind => {
        const editor = editorFor(['    ']);
        if (kind === 'selection') Object.assign(editor.selection, { isEmpty: false });
        if (kind === 'multi') Object.assign(editor, { selections: [editor.selection, editor.selection] });
        if (kind === 'closed') Object.assign(editor.document, { isClosed: true });
        if (kind === 'language') Object.assign(editor.document, { languageId: 'plaintext' });
        if (kind === 'staleCaret') Object.assign(editor.selection, { active: new vscode.Position(8, 4) });
        expect(backspaceNeedsExtension(editor)).toBe(false);
    });
    it('gates the actual keybinding on the narrow cleanup context', () => {
        const manifest = JSON.parse(readFileSync(new URL('../package.json', import.meta.url), 'utf8'));
        const binding = manifest.contributes.keybindings.find((row: {key: string; when?: string}) => row.key === 'backspace' && row.when?.includes(BACKSPACE_NEEDS_EXTENSION_CONTEXT));
        expect(binding.key).toBe('backspace');
        expect(binding.command).toBe('runCommands');
        expect(binding.args.commands).toEqual(['deleteLeft', 'xlide.vba.finishBackspaceCleanup']);
        expect(binding.when).toContain('&& ' + BACKSPACE_NEEDS_EXTENSION_CONTEXT + ' &&');
        expect(binding.when).toContain('!suggestWidgetVisible');
        expect(binding.when).toContain('!inSnippetMode');
    });
    it('updates on editor, selection, text and close without traffic for ordinary typing', () => {
        const ordinary = editorFor(['Sub Demo()', 'ThisWorkbook.Sheets(1).a']);
        activate(ordinary);
        expect(latestValue()).toEqual(['setContext', BACKSPACE_NEEDS_EXTENSION_CONTEXT, false]);
        vi.mocked(vscode.commands.executeCommand).mockClear();
        const changed = vi.mocked(vscode.workspace.onDidChangeTextDocument).mock.calls.at(-1)![0];
        const selection = vi.mocked(vscode.window.onDidChangeTextEditorSelection).mock.calls.at(-1)![0];
        for (let i = 0; i < 40; i++) { changed({ document: ordinary.document } as never); selection({ textEditor: ordinary } as never); }
        expect(vscode.commands.executeCommand).not.toHaveBeenCalled();
        expect(ordinary.document.getText).not.toHaveBeenCalled();
        const indent = editorFor(['    ']);
        Object.assign(vscode.window, { activeTextEditor: indent });
        vi.mocked(vscode.window.onDidChangeActiveTextEditor).mock.calls.at(-1)![0](indent);
        expect(latestValue()).toEqual(['setContext', BACKSPACE_NEEDS_EXTENSION_CONTEXT, true]);
        Object.assign(indent.selection, { active: new vscode.Position(0, 0) });
        selection({ textEditor: indent } as never);
        expect(latestValue()).toEqual(['setContext', BACKSPACE_NEEDS_EXTENSION_CONTEXT, false]);
        Object.assign(indent.selection, { active: new vscode.Position(0, 4) });
        changed({ document: indent.document } as never);
        expect(latestValue()).toEqual(['setContext', BACKSPACE_NEEDS_EXTENSION_CONTEXT, true]);
        Object.assign(indent.document, { isClosed: true });
        vi.mocked(vscode.workspace.onDidCloseTextDocument).mock.calls.at(-1)![0](indent.document);
        expect(latestValue()).toEqual(['setContext', BACKSPACE_NEEDS_EXTENSION_CONTEXT, false]);
    });
});


function cleanupFixture() {
    const lines = ['Sub Demo()', '            ', 'End Sub'];
    const editor = editorFor(lines, 1, 12);
    const remove = vi.fn();
    const edit = vi.fn(async (callback: (builder: { delete: typeof remove }) => void, _options?: { undoStopBefore: boolean; undoStopAfter: boolean }) => {
        callback({ delete: remove }); return true;
    });
    Object.assign(editor, { edit });
    activate(editor);
    const changed = vi.mocked(vscode.workspace.onDidChangeTextDocument).mock.calls.at(-1)![0];
    const selected = vi.mocked(vscode.window.onDidChangeTextEditorSelection).mock.calls.at(-1)![0];
    const finish = vi.mocked(vscode.commands.registerCommand).mock.calls.find(([name]) => name === 'xlide.vba.finishBackspaceCleanup')![1];
    const caret = (character: number) => {
        const position = new vscode.Position(1, character);
        Object.assign(editor.selection, { active: position, anchor: position });
        selected({ textEditor: editor } as never);
    };
    const deleted = (reason?: number) => {
        lines[1] = '        ';
        Object.assign(editor.document, { version: 2 });
        changed({ document: editor.document, reason, contentChanges: [{
            range: new vscode.Range(new vscode.Position(1, 8), new vscode.Position(1, 12)), rangeLength: 4, text: '',
        }] } as never);
    };
    return { editor, lines, edit, remove, finish, changed, caret, deleted };
}

describe('native-first Backspace cleanup lifecycle', () => {
    it('finishes after document, dirty-state notification, command and caret arrive in order', async () => {
        const fixture = cleanupFixture();
        fixture.deleted();
        fixture.changed({ document: fixture.editor.document, contentChanges: [] } as never);
        await fixture.finish();
        expect(fixture.edit).not.toHaveBeenCalled();
        fixture.caret(8);
        expect(fixture.remove).toHaveBeenCalledExactlyOnceWith(new vscode.Range(new vscode.Position(1, 0), new vscode.Position(1, 8)));
        expect(fixture.edit.mock.calls[0][1]).toEqual({ undoStopBefore: false, undoStopAfter: true });
        await fixture.finish();
        expect(fixture.edit).toHaveBeenCalledTimes(1);
        expect(fixture.editor.document.getText).not.toHaveBeenCalled();
    });
    it('retains the preceding caret when its native selection notification arrives first', async () => {
        const fixture = cleanupFixture();
        fixture.caret(8);
        fixture.deleted();
        await fixture.finish();
        expect(fixture.remove).toHaveBeenCalledTimes(1);
    });
    it('can finish immediately when native caret and text already arrived', async () => {
        const fixture = cleanupFixture();
        fixture.deleted(); fixture.caret(8);
        await fixture.finish();
        expect(fixture.edit).toHaveBeenCalledTimes(1);
    });
    it.each(['typing', 'navigation', 'editor', 'closed', 'multi', 'selection', 'version'])('discards cleanup after %s', async kind => {
        const fixture = cleanupFixture();
        fixture.deleted(); fixture.caret(8);
        if (kind === 'typing') {
            fixture.lines[1] += 'a'; Object.assign(fixture.editor.document, { version: 3 });
            fixture.changed({ document: fixture.editor.document, contentChanges: [{
                range: new vscode.Range(new vscode.Position(1, 8), new vscode.Position(1, 8)), rangeLength: 0, text: 'a',
            }] } as never);
        }
        if (kind === 'navigation') fixture.caret(4);
        if (kind === 'editor') Object.assign(vscode.window, { activeTextEditor: editorFor(['other']) });
        if (kind === 'closed') Object.assign(fixture.editor.document, { isClosed: true });
        if (kind === 'multi') Object.assign(fixture.editor, { selections: [fixture.editor.selection, fixture.editor.selection] });
        if (kind === 'selection') Object.assign(fixture.editor.selection, { isEmpty: false });
        if (kind === 'version') Object.assign(fixture.editor.document, { version: 3 });
        await fixture.finish();
        expect(fixture.edit).not.toHaveBeenCalled();
    });
    it('never creates additional deletion for undo or redo events', async () => {
        const fixture = cleanupFixture();
        fixture.deleted(1); fixture.caret(8);
        await fixture.finish();
        expect(fixture.edit).not.toHaveBeenCalled();
    });
    it.each(['rejected', 'thrown'])('never retries or falls back when a cleanup edit is %s', async result => {
        const fixture = cleanupFixture();
        if (result === 'rejected') fixture.edit.mockResolvedValue(false);
        else fixture.edit.mockRejectedValue(new Error('editor closed after dispatch'));
        fixture.deleted(); fixture.caret(8);
        vi.mocked(vscode.commands.executeCommand).mockClear();
        await fixture.finish(); await fixture.finish();
        expect(fixture.edit).toHaveBeenCalledTimes(1);
        expect(vscode.commands.executeCommand).not.toHaveBeenCalled();
    });
});
