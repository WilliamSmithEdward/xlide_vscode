import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type * as vscodeTypes from 'vscode';
vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());
import * as vscode from 'vscode';
import { registerVbaEditorCommands } from '../src/vbaEditorCommands';

function editorFor(comment = false) {
    const lines = comment ? ["' note", "' "] : ['Sub Demo()', 'value'];
    const position = new vscode.Position(1, lines[1].length);
    const selection = { active: position, anchor: position, start: position, end: position, isEmpty: true };
    const document = { uri: vscode.Uri.file('/commands.bas'), languageId: 'vba', version: 1, isClosed: false,
        lineAt: vi.fn((line: number) => ({ text: lines[line] })) };
    return { document, selection, selections: [selection], edit: vi.fn(async () => true) };
}
function select(editor: ReturnType<typeof editorFor>) { Object.assign(vscode.window, { activeTextEditor: editor }); }
function command(name: string) {
    return vi.mocked(vscode.commands.registerCommand).mock.calls.find(call => call[0] === name)![1];
}
function pending<T>() {
    let resolve!: (value: T) => void;
    const promise = new Promise<T>(done => { resolve = done; });
    return { promise, resolve };
}
let context: vscodeTypes.ExtensionContext;
beforeEach(() => {
    vi.mocked(vscode.commands.registerCommand).mockClear();
    vi.mocked(vscode.commands.executeCommand).mockReset();
    context = { subscriptions: [] } as unknown as vscodeTypes.ExtensionContext;
    registerVbaEditorCommands(context);
});
afterEach(() => { context.subscriptions.forEach(item => item.dispose()); vi.restoreAllMocks(); });

describe('editor command lifecycle', () => {
    it.each([['smartBackspace', 'deleteLeft'], ['smartTab', 'tab']])('dispatches ordinary %s without an async comment-check hop', async (name, native) => {
        select(editorFor());
        const result = command('xlide.vba.' + name)();
        expect(vscode.commands.executeCommand).toHaveBeenCalledWith(native);
        await result;
    });

    it.each([['smartBackspace', false], ['smartTab', true]])('does not apply %s fallback to an editor selected while comment clearing waits', async (name, applied) => {
        const first = editorFor(true);
        const edit = pending<boolean>();
        first.edit.mockReturnValueOnce(edit.promise);
        select(first);
        const result = command('xlide.vba.' + name)();
        select(editorFor());
        edit.resolve(applied);
        await result;
        expect(vscode.commands.executeCommand).not.toHaveBeenCalled();
    });

    it('does not delete a later keystroke after the original comment-clearing edit is rejected', async () => {
        const editor = editorFor(true);
        const edit = pending<boolean>();
        editor.edit.mockReturnValueOnce(edit.promise);
        select(editor);
        const result = command('xlide.vba.smartBackspace')();
        editor.document.version++;
        edit.resolve(false);
        await result;
        expect(vscode.commands.executeCommand).not.toHaveBeenCalled();
    });

    it('does not move a newly selected editor after waiting to leave a snippet', async () => {
        select(editorFor());
        const leave = pending<undefined>();
        vi.mocked(vscode.commands.executeCommand).mockImplementationOnce(() => leave.promise);
        const result = command('xlide.vba.leaveSnippetAndCursorMove')('left');
        select(editorFor());
        leave.resolve(undefined);
        await result;
        expect(vscode.commands.executeCommand).toHaveBeenCalledTimes(1);
    });

    it('does not delete at a new caret after a rejected comment-clearing edit', async () => {
        const editor = editorFor(true);
        const edit = pending<boolean>();
        editor.edit.mockReturnValueOnce(edit.promise);
        select(editor);
        const result = command('xlide.vba.smartBackspace')();
        const position = new vscode.Position(0, 1);
        editor.selection = { active: position, anchor: position, start: position, end: position, isEmpty: true };
        edit.resolve(false);
        await result;
        expect(vscode.commands.executeCommand).not.toHaveBeenCalled();
    });

    it('retains native Backspace fallback when a rejected edit is still current', async () => {
        const editor = editorFor(true);
        editor.edit.mockResolvedValueOnce(false);
        select(editor);
        await command('xlide.vba.smartBackspace')();
        expect(vscode.commands.executeCommand).toHaveBeenCalledWith('deleteLeft');
    });

    it('still moves the caret after leaving a snippet in the same editor', async () => {
        select(editorFor());
        await command('xlide.vba.leaveSnippetAndCursorMove')('left');
        expect(vscode.commands.executeCommand).toHaveBeenCalledWith('cursorMove', { to: 'left', by: 'character', value: 1 });
    });

    it('still clears a continued comment on Backspace without an extra native deletion', async () => {
        const editor = editorFor(true);
        select(editor);
        await command('xlide.vba.smartBackspace')();
        expect(editor.edit).toHaveBeenCalledTimes(1);
        expect(vscode.commands.executeCommand).not.toHaveBeenCalled();
    });

    it('still indents after Tab clears the comment in the current editor', async () => {
        const editor = editorFor(true);
        editor.edit.mockImplementationOnce(async () => { editor.document.version++; return true; });
        select(editor);
        await command('xlide.vba.smartTab')();
        expect(vscode.commands.executeCommand).toHaveBeenCalledTimes(1);
    });
});
