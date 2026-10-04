import * as assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { BACKSPACE_NEEDS_EXTENSION_CONTEXT, backspaceNeedsExtension } from '../vbaEditorCommands';
import { runRendererBackspaceProbe } from './rendererBackspaceProbe';
import { activate, closeAllEditors, open, until, workspaceRoot } from './support';

suite('Native Backspace renderer routing', () => {
    const port = Number(process.env.XLIDE_UI_DEBUG_PORT);
    setup(async function () {
        if (!Number.isInteger(port) || port < 1024 || port > 65535) { this.skip(); }
        await activate();
    });
    teardown(async () => {
        if (!port) { return; }
        await vscode.commands.executeCommand('setContext', BACKSPACE_NEEDS_EXTENSION_CONTEXT, false);
        await vscode.commands.executeCommand('workbench.action.revertAndCloseActiveEditor');
        await closeAllEditors();
    });
    for (const routedThroughHost of [false, true]) {
        test(routedThroughHost ? 'control: an extension-bound Backspace waits for the busy host' : 'ordinary Backspace paints its deletion while the extension host is busy', async () => {
            const suffix = routedThroughHost ? 'control' : 'native';
            const file = path.join(workspaceRoot(), `NativeBusyBackspace-${suffix}.bas`);
            fs.writeFileSync(file, 'Sub NativeBusyBackspace()\n    Debug.Print ThisWorkbook.Sheets(1).cez\nEnd Sub\n');
            const document = await open(vscode.Uri.file(file));
            const editor = vscode.window.activeTextEditor!;
            const caret = document.lineAt(1).range.end;
            editor.selection = new vscode.Selection(caret, caret);
            await vscode.commands.executeCommand('hideSuggestWidget');
            await vscode.commands.executeCommand('workbench.action.focusActiveEditorGroup');
            assert.equal(backspaceNeedsExtension(editor), false);
            const result = await runRendererBackspaceProbe('busy', routedThroughHost);
            assert.equal(result.deletedWhileBusy, !routedThroughHost, JSON.stringify(result));
            await until(() => document.lineAt(1).text.endsWith('.ce') || undefined, 'Backspace should eventually reach the extension document');
            console.log(`Renderer Backspace ${suffix}: ${JSON.stringify(result)}`);
        });
    }
});
