import * as assert from 'node:assert/strict';
import * as vscode from 'vscode';
import { activate, closeAllEditors, open, until, writeModule } from './support';

suite('Prefix window surfaces', () => {
    let document: vscode.TextDocument;
    const names = (items: readonly vscode.CompletionItem[]) => items.map(item =>
        typeof item.label === 'string' ? item.label : item.label.label);
    const complete = (offset: number) => vscode.commands.executeCommand<vscode.CompletionList>(
        'vscode.executeCompletionItemProvider', document.uri, document.positionAt(offset));
    suiteSetup(async () => {
        await activate();
        const padding = Array.from({ length: 1200 }, (_, i) =>
            'Sub PrefixPad' + i + '()\r\nEnd Sub\r\n').join('');
        document = await open(await writeModule('PrefixWindowProbe', padding + [
            'Sub PrefixProbe()', ...Array.from({ length: 6 }, () => '    Debug.Print ThisWorkbook.Sheets(1).Name'),
            '    Debug.Print ThisWorkbook.Sheets(1). _', '        Name',
            '    x = 1: On Error ', 'End Sub', '',
        ].join('\r\n')));
    });
    suiteTeardown(async () => {
        if (document?.isDirty) { await document.save(); }
        await closeAllEditors();
    });

    test('moving across late member chains retains completion and hover results', async () => {
        const source = document.getText();
        const offsets = [...source.matchAll(/Sheets\(1\)\.Name/g)].map(match => match.index! + 'Sheets(1).'.length);
        await until(async () => names((await complete(offsets[0]))?.items ?? []).includes('Name') ? true : undefined,
            'late dot completion should include Name');
        for (const offset of offsets) {
            assert.ok(names((await complete(offset))?.items ?? []).includes('Name'));
            const hover = await vscode.commands.executeCommand<vscode.Hover[]>('vscode.executeHoverProvider',
                document.uri, document.positionAt(offset + 2));
            const text = (hover ?? []).flatMap(item => item.contents.map(content =>
                typeof content === 'string' ? content : content.value)).join('\n');
            assert.match(text, /Name/);
        }
        const keyword = source.indexOf('x = 1: On Error ') + 'x = 1: On Error '.length;
        assert.ok(names((await complete(keyword))?.items ?? []).includes('Resume Next'));
    });

    test('each typed character updates completion on a continued member chain', async () => {
        const source = document.getText();
        const start = source.lastIndexOf('        Name') + '        '.length;
        const editor = vscode.window.activeTextEditor!;
        assert.ok(await editor.edit(edit => edit.delete(new vscode.Range(document.positionAt(start), document.positionAt(start + 4)))));
        let offset = start;
        for (const character of 'Name') {
            editor.selection = new vscode.Selection(document.positionAt(offset), document.positionAt(offset));
            assert.ok(await editor.edit(edit => edit.insert(document.positionAt(offset), character)));
            offset += 1;
            editor.selection = new vscode.Selection(document.positionAt(offset), document.positionAt(offset));
            const result = await complete(offset);
            assert.ok(names(result?.items ?? []).includes('Name'), 'Name should remain available after ' + character);
            await new Promise(resolve => setTimeout(resolve, 30));
        }
    });
});
