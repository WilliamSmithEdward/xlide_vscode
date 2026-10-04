import * as assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { activate, closeAllEditors, open, until, workspaceRoot } from './support';
import { runRendererCursorProbe, type CursorDirection } from './rendererCursorProbe';

suite('Native snippet renderer routing', () => {
    const port = Number(process.env.XLIDE_UI_DEBUG_PORT);
    setup(async function () {
        if (!Number.isInteger(port) || port < 1024 || port > 65535) { this.skip(); }
        await activate();
    });
    teardown(async () => {
        if (!port) { return; }
        await vscode.commands.executeCommand('leaveSnippet');
        await vscode.commands.executeCommand('workbench.action.revertAndCloseActiveEditor');
        await closeAllEditors();
    });
    async function fixture(snippet: boolean): Promise<vscode.TextEditor> {
        const file = path.join(workspaceRoot(), 'NativeSnippetRouting.bas');
        fs.writeFileSync(file, 'Sub NativeSnippetProbe()\n    Dim latency As Long\n    BeforeLine\n    MiddleMarker\n    AfterLine\nEnd Sub\n');
        await open(vscode.Uri.file(file));
        const editor = vscode.window.activeTextEditor!;
        if (snippet) {
            assert.equal(await editor.insertSnippet(new vscode.SnippetString('${1:MiddleMarker}'), new vscode.Range(3, 4, 3, 16)), true);
        }
        const start = new vscode.Position(3, 10);
        editor.selection = new vscode.Selection(start, start);
        await vscode.commands.executeCommand('hideSuggestWidget');
        await vscode.commands.executeCommand('workbench.action.focusActiveEditorGroup');
        return editor;
    }
    for (const snippet of [false, true]) {
        for (const direction of ['up', 'down', 'left', 'right'] as const) {
            test(`${snippet ? 'snippet' : 'ordinary'} ${direction} moves while the host is busy`, async () => {
                const editor = await fixture(snippet);
                const result = await runRendererCursorProbe(direction, snippet);
                console.log('Renderer native arrow:', JSON.stringify({ snippet, direction, ...result }));
                const expected = expectedCaret(direction);
                await until(() => editor.selection.active.isEqual(expected) || undefined, 'physical arrow must move exactly once');
                assert.equal(result.movedWhileBusy, true, JSON.stringify(result));
                assert.equal(result.after.placeholders, 0);
            });
        }
        test(`${snippet ? 'snippet' : 'ordinary'} arrow then typing edits the intended row while the host is busy`, async () => {
            const editor = await fixture(snippet);
            const result = await runRendererCursorProbe('up', snippet, true);
            console.log('Renderer arrow then typing:', JSON.stringify({ snippet, ...result }));
            assert.equal(result.typedWhileBusy, true, JSON.stringify(result));
            await until(() => editor.document.lineAt(2).text === '    BeforezLine' || undefined, 'rapid typing must reach the row selected by the arrow');
            assert.equal(editor.document.lineAt(3).text, '    MiddleMarker');
            assert.equal(result.after.placeholders, 0);
            assert.equal(result.after.middleTyped, false);
        });
    }
});

function expectedCaret(direction: CursorDirection): vscode.Position {
    return new vscode.Position(direction === 'up' ? 2 : direction === 'down' ? 4 : 3,
        direction === 'left' ? 9 : direction === 'right' ? 11 : 10);
}
