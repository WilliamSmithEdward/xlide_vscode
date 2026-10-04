import * as assert from 'node:assert/strict';
import * as vscode from 'vscode';
import { activate, closeAllEditors, open, until, writeModule } from './support';

suite('Identifier snapshot surfaces', () => {
    let document: vscode.TextDocument;
    const label = (item: vscode.CompletionItem) => typeof item.label === 'string' ? item.label : item.label.label;
    const complete = (line: number) => vscode.commands.executeCommand<vscode.CompletionList>(
        'vscode.executeCompletionItemProvider', document.uri, new vscode.Position(line, 10));

    suiteSetup(async () => {
        await activate();
        const source = [
            'Option Explicit', 'Sub SnapshotFirst()', '    Dim CachedValue As Long',
            '    cached', '    ThisWorkbook.Sheets(1).Name = "snapshot"', 'End Sub',
            'Sub SnapshotSecond()', '    Dim CachedOther As String', '    cached', 'End Sub',
        ].join('\r\n') + Array.from({ length: 1200 }, (_, index) =>
            '\r\nSub SnapshotExtra' + index + '()\r\nEnd Sub\r\n').join('');
        document = await open(await writeModule('IdentifierSnapshotProbe', source));
    });
    suiteTeardown(async () => {
        if (document?.isDirty) { await document.save(); }
        await closeAllEditors();
    });

    test('repeated completion selects the current procedure in a large module', async () => {
        await until(async () => (await complete(3))?.items.some(item => label(item) === 'CachedValue') ? true : undefined,
            'first procedure completion should contain its local');
        for (let attempt = 0; attempt < 6; attempt++) {
            const first = (await complete(3))?.items.map(label) ?? [];
            const second = (await complete(8))?.items.map(label) ?? [];
            assert.ok(first.includes('CachedValue'));
            assert.ok(!first.includes('CachedOther'));
            assert.ok(second.includes('CachedOther'));
            assert.ok(!second.includes('CachedValue'));
        }
    });

    test('declaration edits refresh completion and hover without breaking chained dot completion', async () => {
        const editor = vscode.window.activeTextEditor!;
        assert.ok(await editor.edit(edit => edit.replace(document.lineAt(2).range, '    Dim CachedRenamed As String')));
        await until(async () => {
            const items = (await complete(3))?.items ?? [];
            return items.some(item => label(item) === 'CachedRenamed') && !items.some(item => label(item) === 'CachedValue')
                ? true : undefined;
        }, 'edited declaration should replace the cached local');
        const hover = await vscode.commands.executeCommand<vscode.Hover[]>('vscode.executeHoverProvider',
            document.uri, new vscode.Position(2, 12));
        assert.ok(hover?.some(item => item.contents.some(content =>
            typeof content === 'string' ? /CachedRenamed.*String/s.test(content)
                : /CachedRenamed.*String/s.test(content.value))), 'hover should show the new declaration and type');
        const dot = await vscode.commands.executeCommand<vscode.CompletionList>('vscode.executeCompletionItemProvider',
            document.uri, new vscode.Position(4, '    ThisWorkbook.Sheets(1).'.length), '.');
        assert.ok(dot?.items.some(item => label(item) === 'Name'));
    });
});
