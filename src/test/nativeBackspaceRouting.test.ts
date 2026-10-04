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
    test('typing out of a cleanup line keeps Backspace native during a busy host', async () => {
        const file = path.join(workspaceRoot(), 'NativeBusyTransition.bas');
        fs.writeFileSync(file, 'Sub NativeBusyBackspace()\n            \nEnd Sub\n');
        const document = await open(vscode.Uri.file(file));
        const editor = vscode.window.activeTextEditor!;
        const caret = document.lineAt(1).range.end;
        editor.selection = new vscode.Selection(caret, caret);
        await vscode.commands.executeCommand('hideSuggestWidget');
        await vscode.commands.executeCommand('workbench.action.focusActiveEditorGroup');
        assert.equal(backspaceNeedsExtension(editor), true);
        // This is the real state of a blank indented line, acknowledged in
        // the renderer before the host becomes busy and typing enters code.
        await vscode.commands.executeCommand('setContext', BACKSPACE_NEEDS_EXTENSION_CONTEXT, true);
        const result = await runRendererBackspaceProbe('transition');
        console.log('Renderer Backspace cleanup transition:', JSON.stringify(result));
        assert.equal(result.deletedWhileBusy, true, JSON.stringify(result));
        await until(() => document.lineAt(1).text.endsWith('.ce') || undefined, 'transition Backspace should reach the extension document');
    });

    for (const fixture of [
        { name: 'Indent', source: 'Sub Demo()\n            \nEnd Sub\n', line: 1, before: '            ', after: '' },
        { name: 'Comment', source: "Sub Demo()\n    'note\n    ' \nEnd Sub\n", line: 2, before: "    ' ", after: '    ' },
        { name: 'LongComment', source: "Sub Demo()\n    ''''note\n    '''' \nEnd Sub\n", line: 2, before: "    '''' ", after: '    ' },
    ]) {
        test(`native-first ${fixture.name} cleanup is one undoable keyboard action`, async () => {
            const file = path.join(workspaceRoot(), `NativeCleanup${fixture.name}.bas`);
            fs.writeFileSync(file, fixture.source);
            const document = await open(vscode.Uri.file(file));
            const editor = vscode.window.activeTextEditor!;
            const caret = document.lineAt(fixture.line).range.end;
            editor.selection = new vscode.Selection(caret, caret);
            await vscode.commands.executeCommand('hideSuggestWidget');
            await vscode.commands.executeCommand('workbench.action.focusActiveEditorGroup');
            assert.equal(backspaceNeedsExtension(editor), true);
            await vscode.commands.executeCommand('setContext', BACKSPACE_NEEDS_EXTENSION_CONTEXT, true);
            const events: unknown[] = [];
            const subscriptions = [
                vscode.workspace.onDidChangeTextDocument(event => {
                    if (event.document === document) { events.push({ kind: 'text', version: document.version,
                        caret: editor.selection.active, length: document.lineAt(fixture.line).text.length,
                        changes: event.contentChanges.map(change => ({ start: change.range.start.character, end: change.range.end.character })) }); }
                }),
                vscode.window.onDidChangeTextEditorSelection(event => {
                    if (event.textEditor === editor) { events.push({ kind: 'selection', version: document.version,
                        caret: editor.selection.active, length: document.lineAt(fixture.line).text.length }); }
                }),
            ];
            let result;
            try {
                result = await runRendererBackspaceProbe('cleanup', false, {
                    line: fixture.line, before: fixture.before, after: [fixture.after],
                });
            } catch (error) { throw new Error(`${String(error)}; events=${JSON.stringify(events)}`); }
            finally { subscriptions.forEach(subscription => subscription.dispose()); }
            assert.equal(document.lineAt(fixture.line).text, fixture.after);
            await vscode.commands.executeCommand('undo');
            await until(() => document.getText() === fixture.source || undefined,
                'one undo should restore both the native deletion and remaining cleanup');
            console.log(`Renderer cleanup ${fixture.name}: ${JSON.stringify(result)}`);
        });
    }

    for (const routedThroughHost of [false, true]) {
        test(routedThroughHost ? 'a stale cleanup context keeps ordinary Backspace native during a busy host' : 'ordinary Backspace paints its deletion while the extension host is busy', async () => {
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
            assert.equal(result.deletedWhileBusy, true, JSON.stringify(result));
            await until(() => document.lineAt(1).text.endsWith('.ce') || undefined, 'Backspace should eventually reach the extension document');
            console.log(`Renderer Backspace ${suffix}: ${JSON.stringify(result)}`);
        });
    }
});
