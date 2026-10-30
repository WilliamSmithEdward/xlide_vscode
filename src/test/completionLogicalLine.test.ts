import * as assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { encodeModuleUri } from '../xlideFileSystem';
import { writeModule } from '../vba/projectService';
import { activate, EXTENSION_ID, open, until, workspaceRoot } from './support';

// Issue #1527: a line break ends a VBA statement, so completion reads only the
// caret's logical line. The VBE reports `Dim t As`, `Set c = New`, `Dim` and
// `t =` with an empty row below as a syntax error on that line in all four
// hosts, and the row below is a fresh statement.
suite('Completion reads the caret logical line', () => {
    const label = (item: vscode.CompletionItem) => typeof item.label === 'string' ? item.label : item.label.label;
    const lines = [
        'Option Explicit',
        'Sub Probe()',
        'Dim t As',
        '',
        'Set c = New',
        '',
        't =',
        '',
        'Dim u As Long ',
        'Dim v As _',
        '    Lo',
        'End Sub',
        '',
    ];
    const at = (row: number, column?: number) => new vscode.Position(row, column ?? lines[row].length);

    suiteSetup(async () => { await activate(); });

    for (const [host, extension] of [['Excel', 'xlsm'], ['Word', 'docm'], ['PowerPoint', 'pptm'], ['Access', 'accdb']]) {
        test(`${host}: the row below an unfinished line is a statement start, and a finished type name stays closed`, async () => {
            const installed = vscode.extensions.getExtension(EXTENSION_ID)!;
            const file = path.join(workspaceRoot(), `LogicalLine.${extension}`);
            fs.copyFileSync(path.join(installed.extensionUri.fsPath, 'assets/templates', `blank.${extension}`), file);
            writeModule(file, 'LogicalLineProbe', lines.join('\r\n'));
            const document = await open(encodeModuleUri(file, 'LogicalLineProbe'));
            const timings: number[] = [];
            const complete = async (position: vscode.Position) => {
                const before = performance.now();
                const list = await vscode.commands.executeCommand<vscode.CompletionList>(
                    'vscode.executeCompletionItemProvider', document.uri, position);
                timings.push(performance.now() - before);
                return (list?.items ?? []).map(label);
            };
            // The continued declaration is still a type position.
            await until(async () => (await complete(at(10))).includes('Long') || undefined,
                `${host}: a continued As should offer Long`);
            for (const row of [3, 5, 7]) {
                const names = await complete(at(row));
                assert.ok(names.includes('DoEvents'), `${host} row ${row}: statement start should offer DoEvents`);
                assert.ok(!names.includes('LongPtr'), `${host} row ${row}: no type list below the line`);
                assert.ok(!names.includes('True'), `${host} row ${row}: no value list below the line`);
            }
            assert.ok(!(await complete(at(8))).includes('LongLong'), `${host}: a finished type name stays closed`);
            const cold = timings[0];
            const warm: number[] = [];
            for (let i = 0; i < 10; i++) {
                for (const row of [3, 5, 7, 8]) {
                    const before = performance.now();
                    await vscode.commands.executeCommand('vscode.executeCompletionItemProvider', document.uri, at(row));
                    warm.push(performance.now() - before);
                }
            }
            warm.sort((a, b) => a - b);
            process.stdout.write(`#1527 ${host} provider ms: first ${cold.toFixed(2)}, warm median ${warm[Math.floor(warm.length / 2)].toFixed(2)}, ` +
                `p95 ${warm[Math.floor(warm.length * 0.95)].toFixed(2)}, max ${warm[warm.length - 1].toFixed(2)}\n`);
            await vscode.commands.executeCommand('workbench.action.closeActiveEditor');
        });
    }
});
