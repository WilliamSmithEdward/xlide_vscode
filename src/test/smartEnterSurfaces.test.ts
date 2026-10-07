import * as assert from 'node:assert/strict';
import * as vscode from 'vscode';
import { activate, closeAllEditors, open, until, writeModule } from './support';

suite('Smart Enter surfaces', () => {
    suiteSetup(activate);
    teardown(async () => {
        for (const document of vscode.workspace.textDocuments) {
            if (document.isDirty && document.uri.scheme === 'xlide-vba') { await document.save(); }
        }
        await closeAllEditors();
    });

    async function enter(moduleName: string, source: string, line: number, indentation: string) {
        const document = await open(await writeModule(moduleName, source));
        const editor = vscode.window.activeTextEditor!;
        const column = document.lineAt(line).text.length;
        editor.selection = new vscode.Selection(line, column, line, column);
        const start = performance.now();
        assert.ok(await editor.edit(edit => edit.insert(new vscode.Position(line, column), '\r\n' + indentation),
            { undoStopBefore: false, undoStopAfter: false }));
        console.log(moduleName, 'Enter edit command (ms):', performance.now() - start);
        return document;
    }

    test('ordinary Enter in a large module leaves the new body line unchanged', async () => {
        const procedures = Array.from({ length: 1200 }, (_, index) =>
            'Sub Probe' + index + '()\r\n    Debug.Print "probe"\r\nEnd Sub\r\n').join('\r\n');
        const source = procedures + '\r\nSub ActiveProbe()\r\n    value = 1\r\nEnd Sub\r\n';
        const line = source.split('\r\n').length - 3;
        const document = await enter('EnterLarge', source, line, '    ');
        await new Promise(resolve => setTimeout(resolve, 800));
        assert.equal(document.lineAt(line + 1).text, '    ');
        assert.equal(document.lineAt(line + 2).text, 'End Sub');
    });

    test('comment Enter preserves indentation and apostrophe spacing', async () => {
        const document = await enter('EnterComment', "Sub T()\r\n    '''  note\r\nEnd Sub\r\n", 1, '    ');
        await until(() => document.lineAt(2).text === "    '''  " ? true : undefined,
            'comment continuation should preserve the prefix');
    });

    test('With member Enter seeds a dot inside the open block', async () => {
        const document = await enter('EnterWith',
            'Sub T()\r\n    With ActiveSheet\r\n        .Name = "probe"\r\n    End With\r\nEnd Sub\r\n', 2, '        ');
        await until(() => document.lineAt(3).text === '        .' ? true : undefined,
            'With continuation should seed a dot');
        const completions = await vscode.commands.executeCommand<vscode.CompletionList>(
            'vscode.executeCompletionItemProvider', document.uri, new vscode.Position(3, 9), '.');
        assert.ok(completions?.items.some(item => (typeof item.label === 'string' ? item.label : item.label.label) === 'Name'));
    });

    test('block Enter inserts the missing closer', async () => {
        const document = await enter('EnterBlock', 'Sub T()\r\n    If ready Then\r\nEnd Sub\r\n', 1, '    ');
        await until(() => /End If/i.test(document.getText()) ? true : undefined,
            'If continuation should insert End If');
        assert.equal((document.getText().match(/End If/gi) ?? []).length, 1);
    });
});
