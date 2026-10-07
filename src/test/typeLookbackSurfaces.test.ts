import * as assert from 'node:assert/strict';
import * as vscode from 'vscode';
import { activate, closeAllEditors, insertTypedCharacter, open, until, writeModule } from './support';

suite('Type lookback surfaces', () => {
    let document: vscode.TextDocument;
    const label = (item: vscode.CompletionItem) => typeof item.label === 'string' ? item.label : item.label.label;

    suiteSetup(async () => {
        await activate();
        const padding = Array.from({ length: 1200 }, (_, index) =>
            'Sub TypePadding' + index + '()\r\nEnd Sub\r\n').join('');
        document = await open(await writeModule('TypeLookbackProbe', padding +
            'Sub TypeProbe()\r\n    Dim item As _\r\n        Excel.Wor\r\n    Dim CanonicalValue As Long\r\n    \r\nEnd Sub\r\n'));
    });
    suiteTeardown(async () => {
        if (document?.isDirty) { await document.save(); }
        await closeAllEditors();
    });

    test('qualified type completion survives prefix edits after a continuation in a large module', async () => {
        const start = document.getText().lastIndexOf('Excel.Wor') + 'Excel.'.length;
        const position = document.positionAt(start + 3);
        const complete = () => vscode.commands.executeCommand<vscode.CompletionList>(
            'vscode.executeCompletionItemProvider', document.uri, position);
        await until(async () => (await complete())?.items.some(item => label(item) === 'Worksheet') ? true : undefined,
            'continued qualified type should offer Worksheet');
        const editor = vscode.window.activeTextEditor!;
        assert.ok(await editor.edit(edit => edit.replace(new vscode.Range(document.positionAt(start), position), 'Ran')));
        const next = await complete();
        assert.ok(next?.items.some(item => label(item) === 'Range'));
        assert.ok(!next?.items.some(item => label(item) === 'Worksheet'));
    });

    test('typing and saving still recase repeated variables near the end of a large module', async () => {
        await document.save();
        const editor = vscode.window.activeTextEditor!;
        const line = document.lineCount - 3;
        for (const character of 'canonicalvalue = canonicalvalue + canonicalvalue') {
            await new Promise(resolve => setTimeout(resolve, 30));
            await insertTypedCharacter(document, editor, line, character, { undoStopBefore: false, undoStopAfter: false });
        }
        await document.save();
        await until(() => document.lineAt(line).text.includes('CanonicalValue = CanonicalValue + CanonicalValue')
            ? true : undefined, 'typed line should retain canonical casing at save');
    });
});
