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
        document: { languageId: 'vba', uri: vscode.Uri.file('/backspace.bas'), isClosed: false, lineCount: lines.length,
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
beforeEach(() => { vi.mocked(vscode.commands.executeCommand).mockClear(); vi.mocked(vscode.window.onDidChangeActiveTextEditor).mockClear(); vi.mocked(vscode.window.onDidChangeTextEditorSelection).mockClear(); vi.mocked(vscode.workspace.onDidChangeTextDocument).mockClear(); vi.mocked(vscode.workspace.onDidCloseTextDocument).mockClear(); });
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
        const binding = manifest.contributes.keybindings.find((row: {command: string}) => row.command === 'xlide.vba.smartBackspace');
        expect(binding.key).toBe('backspace');
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
