import * as assert from 'node:assert/strict';
import * as vscode from 'vscode';
import { activate, closeAllEditors, insertTypedCharacter, open, until, writeModule } from './support';

suite('Immediate completion surfaces', () => {
    let document: vscode.TextDocument;
    let bodyLine: number;
    const label = (item: vscode.CompletionItem) => typeof item.label === 'string' ? item.label : item.label.label;
    const complete = (position: vscode.Position) => vscode.commands.executeCommand<vscode.CompletionList>(
        'vscode.executeCompletionItemProvider', document.uri, position);

    suiteSetup(async () => {
        await activate();
        await writeModule('ImmediateApi', 'Public Sub InstantExternalApi()\r\nEnd Sub\r\n');
        const padding = Array.from({ length: 1200 }, (_, index) =>
            'Sub ImmediatePad' + index + '()\r\nEnd Sub\r\n').join('');
        document = await open(await writeModule('ImmediateProbe', padding +
            'Sub Prompt()\r\n    Dim InstantValue As Long\r\n    Instant\r\n    ThisWorkbook.Sheets(1).\r\nEnd Sub\r\n'));
        bodyLine = document.lineCount - 4;
        await until(async () => (await complete(document.lineAt(bodyLine).range.end))?.items.some(item => label(item) === 'InstantExternalApi')
            ? true : undefined, 'project should be loaded before measuring warm typing');
        assert.ok(await vscode.window.activeTextEditor!.edit(edit => edit.replace(document.lineAt(bodyLine).range, '    ')));
        await document.save();
    });
    suiteTeardown(async () => {
        if (document?.isDirty) { await document.save(); }
        await closeAllEditors();
    });

    test('updates local completion after each actual keystroke and records response latency', async () => {
        const editor = vscode.window.activeTextEditor!;
        const samples: number[] = [];
        const editAttempts: number[] = [];
        for (const character of 'instantvalue') {
            await new Promise(resolve => setTimeout(resolve, 30));
            const start = performance.now();
            editAttempts.push(await insertTypedCharacter(document, editor, bodyLine, character));
            const result = await complete(document.lineAt(bodyLine).range.end);
            samples.push(performance.now() - start);
            assert.ok(result?.items.some(item => label(item) === 'InstantValue'), 'each prefix should retain the local candidate');
        }
        samples.sort((a, b) => a - b);
        console.log('Actual edit-to-completion-result latency (ms):', JSON.stringify({
            editAttempts, median: samples[Math.floor(samples.length / 2)], p95: samples[Math.ceil(samples.length * 0.95) - 1], samples: samples.length, sortedSamples: samples,
        }));
        const dotLine = bodyLine + 1;
        for (const character of 'name') {
            await new Promise(resolve => setTimeout(resolve, 30));
            await insertTypedCharacter(document, editor, dotLine, character);
            assert.ok((await complete(document.lineAt(dotLine).range.end))?.items.some(item => label(item) === 'Name'));
        }
    });

    test('cross-module candidates become available after background warming', async () => {
        const editor = vscode.window.activeTextEditor!;
        assert.ok(await editor.edit(edit => edit.replace(document.lineAt(bodyLine).range, '    InstantExternal')));
        await until(async () => (await complete(document.lineAt(bodyLine).range.end))?.items.some(item => label(item) === 'InstantExternalApi')
            ? true : undefined, 'background project facts should populate subsequent completion requests');
    });
});
