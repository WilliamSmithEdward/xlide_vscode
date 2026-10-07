import * as assert from 'node:assert/strict';
import * as vscode from 'vscode';
import { activate, closeAllEditors, open, until, writeModule } from './support';

suite('Identifier surfaces', () => {
    suiteSetup(activate);
    teardown(async () => {
        for (const document of vscode.workspace.textDocuments) {
            if (document.isDirty && document.uri.scheme === 'xlide-vba') { await document.save(); }
        }
        await closeAllEditors();
    });

    function name(item: vscode.CompletionItem): string {
        return typeof item.label === 'string' ? item.label : item.label.label;
    }
    function documentation(item: vscode.CompletionItem): string {
        return typeof item.documentation === 'string' ? item.documentation : item.documentation?.value ?? '';
    }

    test('prefix-filtered completion retains project documentation and automatic casing in a large module', async () => {
        await writeModule('IdentifierApi', [
            "''' <summary>Runs the named operation.</summary>",
            'Public Sub ZzTargetOperation(ByVal amount As Long)', 'End Sub', '',
        ].join('\r\n'));
        const source = Array.from({ length: 1200 }, (_, index) => [
            "''' <summary>Probe documentation.</summary>",
            'Sub Probe' + index + '(ByVal value As Long)', 'End Sub', '',
        ].join('\r\n')).join('\r\n') + [
            '', 'Sub ActiveProbe()', '    Dim ZzValue As Long', '    zz', 'End Sub', '',
        ].join('\r\n');
        const document = await open(await writeModule('IdentifierLarge', source));
        const line = source.split('\r\n').findIndex(text => text === '    zz');
        const position = new vscode.Position(line, 6);
        const list = await until(async () => {
            const items = await vscode.commands.executeCommand<vscode.CompletionList>(
                'vscode.executeCompletionItemProvider', document.uri, position);
            return items?.items.some(item => name(item) === 'ZzTargetOperation') ? items : undefined;
        }, 'project procedure should appear alongside the matching local');
        assert.ok(list.items.some(item => name(item) === 'ZzValue'));
        assert.ok(!list.items.some(item => /^Probe\d+$/.test(name(item))));
        const procedure = list.items.find(item => name(item) === 'ZzTargetOperation')!;
        assert.ok(procedure.detail?.includes('amount As Long'));
        assert.ok(documentation(procedure).includes('Runs the named operation.'));

        const editor = vscode.window.activeTextEditor!;
        editor.selection = new vscode.Selection(line, 6, line, 6);
        let column = 6;
        for (const character of 'value') {
            assert.ok(await editor.edit(edit => edit.insert(new vscode.Position(line, column), character),
                { undoStopBefore: false, undoStopAfter: false }));
            column++;
            editor.selection = new vscode.Selection(line, column, line, column);
            await new Promise(resolve => setTimeout(resolve, 30));
        }
        await until(() => document.lineAt(line).text === '    ZzValue' ? true : undefined,
            'idle casing should restore the local declaration spelling');
    });

    test('a local procedure keeps its own documentation over a project duplicate', async () => {
        await writeModule('IdentifierDuplicate', [
            "''' <summary>Remote documentation.</summary>",
            'Public Sub ZzShadowProbe(ByVal remoteValue As String)', 'End Sub', '',
        ].join('\r\n'));
        const source = [
            "''' <summary>Local documentation.</summary>",
            'Private Sub ZzShadowProbe(ByVal localValue As Long)', 'End Sub',
            'Sub Caller()', '    zzshadowp', 'End Sub', '',
        ].join('\r\n');
        const document = await open(await writeModule('IdentifierShadow', source));
        const list = await until(async () => {
            const result = await vscode.commands.executeCommand<vscode.CompletionList>(
                'vscode.executeCompletionItemProvider', document.uri, new vscode.Position(4, 13));
            return result?.items.some(item => name(item) === 'ZzShadowProbe') ? result : undefined;
        }, 'local procedure should be offered');
        const matches = list.items.filter(item => name(item) === 'ZzShadowProbe');
        assert.equal(matches.length, 1);
        assert.ok(documentation(matches[0]).includes('Local documentation.'));
        assert.ok(!documentation(matches[0]).includes('Remote documentation.'));
        assert.ok(matches[0].detail?.includes('localValue As Long'));
    });
});
