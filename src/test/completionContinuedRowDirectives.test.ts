import * as assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { encodeModuleUri } from '../xlideFileSystem';
import { writeModule } from '../vba/projectService';
import { activate, EXTENSION_ID, open, until, workspaceRoot } from './support';

// A test directive is a comment line of its own. A row that carries on a
// statement from the row above cannot hold one, so completion leaves the
// directives out there. A row that carries on a comment is still comment text.
suite('Completion on a continued row', () => {
    const label = (item: vscode.CompletionItem) => typeof item.label === 'string' ? item.label : item.label.label;
    const lines = [
        'Option Explicit',
        'Sub Probe()',
        'Dim t As Long _',
        '    ',
        'Dim total As Long',
        'total = 1 + _',
        '    ',
        "' Checks the total _",
        '    ',
        'End Sub',
        '',
    ];
    const at = (row: number) => new vscode.Position(row, lines[row].length);

    suiteSetup(async () => { await activate(); });

    for (const [host, extension] of [['Excel', 'xlsm'], ['Word', 'docm'], ['PowerPoint', 'pptm'], ['Access', 'accdb']]) {
        test(`${host}: no test directives on a row that continues a statement`, async () => {
            const installed = vscode.extensions.getExtension(EXTENSION_ID)!;
            const file = path.join(workspaceRoot(), `ContinuedRow.${extension}`);
            fs.copyFileSync(path.join(installed.extensionUri.fsPath, 'assets/templates', `blank.${extension}`), file);
            writeModule(file, 'ContinuedRowProbe', lines.join('\r\n'));
            const document = await open(encodeModuleUri(file, 'ContinuedRowProbe'));
            const complete = async (position: vscode.Position) => {
                const list = await vscode.commands.executeCommand<vscode.CompletionList>(
                    'vscode.executeCompletionItemProvider', document.uri, position);
                return (list?.items ?? []).map(label);
            };
            await until(async () => (await complete(at(6))).includes('total') || undefined,
                `${host}: the continued expression should offer total`);
            const directives = ['@xlide-test', '@xlide-test-skip', '@xlide-test-xfail'];
            for (const row of [3, 6]) {
                const names = await complete(at(row));
                assert.deepEqual(directives.filter(name => names.includes(name)), [], `${host} row ${row}: no test directives`);
            }
            const commentRow = await complete(at(8));
            assert.deepEqual(directives.filter(name => commentRow.includes(name)), directives,
                `${host}: a continued comment keeps the directives`);
            await vscode.commands.executeCommand('workbench.action.closeActiveEditor');
        });
    }
});
