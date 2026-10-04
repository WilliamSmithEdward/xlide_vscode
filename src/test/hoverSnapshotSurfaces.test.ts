import * as assert from 'node:assert/strict';
import * as vscode from 'vscode';
import { activate, closeAllEditors, open, until, writeModule } from './support';

suite('Hover snapshot surfaces', () => {
    let document: vscode.TextDocument;
    const hoverText = async (offset: number) => {
        const hovers = await vscode.commands.executeCommand<vscode.Hover[]>('vscode.executeHoverProvider',
            document.uri, document.positionAt(offset));
        return (hovers ?? []).flatMap(item => item.contents.map(content =>
            typeof content === 'string' ? content : content.value)).join('\n');
    };

    suiteSetup(async () => {
        await activate();
        const padding = Array.from({ length: 1200 }, (_, index) =>
            'Sub HoverPad' + index + '()\r\nEnd Sub\r\n').join('');
        document = await open(await writeModule('HoverSnapshotProbe', padding + [
            'Sub HoverFirst()', '    Dim HoverValue As Long', '    Debug.Print HoverValue',
            '    Debug.Print ActiveWorkbook', 'End Sub',
            'Sub HoverSecond()', '    Dim HoverValue As String', '    Debug.Print HoverValue',
            '    ThisWorkbook.Sheets(1).Name = "hover"', 'End Sub', '',
        ].join('\r\n')));
    });
    suiteTeardown(async () => {
        if (document?.isDirty) { await document.save(); }
        await closeAllEditors();
    });

    test('repeated hover at different symbols retains procedure scope in a large module', async () => {
        const source = document.getText();
        const first = source.indexOf('Debug.Print HoverValue') + 'Debug.Print '.length + 2;
        const second = source.lastIndexOf('Debug.Print HoverValue') + 'Debug.Print '.length + 2;
        const host = source.indexOf('Debug.Print ActiveWorkbook') + 'Debug.Print '.length + 2;
        await until(async () => /HoverValue As Long/.test(await hoverText(first)) ? true : undefined,
            'first local hover should resolve');
        for (let attempt = 0; attempt < 5; attempt++) {
            assert.match(await hoverText(first), /HoverValue As Long/);
            assert.match(await hoverText(second), /HoverValue As String/);
            assert.match(await hoverText(host), /ActiveWorkbook As Workbook/);
        }
    });

    test('source edits refresh local hover and host-global shadowing while other scopes keep dot completion', async () => {
        const source = document.getText();
        const start = source.indexOf('    Dim HoverValue As Long');
        const end = start + '    Dim HoverValue As Long'.length;
        const editor = vscode.window.activeTextEditor!;
        assert.ok(await editor.edit(edit => edit.replace(new vscode.Range(document.positionAt(start), document.positionAt(end)),
            '    Dim HoverValue As Double\r\n    Dim ActiveWorkbook As String')));
        const changed = document.getText();
        const local = changed.indexOf('Debug.Print HoverValue') + 'Debug.Print '.length + 2;
        const host = changed.indexOf('Debug.Print ActiveWorkbook') + 'Debug.Print '.length + 2;
        assert.match(await hoverText(local), /HoverValue As Double/);
        assert.match(await hoverText(host), /ActiveWorkbook As String/);
        const dot = changed.lastIndexOf('ThisWorkbook.Sheets(1).') + 'ThisWorkbook.Sheets(1).'.length;
        const complete = await vscode.commands.executeCommand<vscode.CompletionList>('vscode.executeCompletionItemProvider',
            document.uri, document.positionAt(dot), '.');
        assert.ok(complete?.items.some(item => (typeof item.label === 'string' ? item.label : item.label.label) === 'Name'));
    });
});
