// Issue #1240: exercise workbook reference parsing and the actual editor providers.
import * as assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { openMacroContainer } from '../vba/macroContainer';
import { VbaProject } from '../vba/vbaProject';
import { buildRegisteredReference } from '../vba/vbaProjectReferences';
import { writeModule, listReferences } from '../vba/projectService';
import { encodeModuleUri } from '../xlideFileSystem';
import { activate, closeAllEditors, open, until, workbookPath, workspaceRoot } from './support';

const MODULE = 'ReferenceProbe';
const SOURCE = [
    'Option Explicit',
    'Public Sub Probe()',
    '    Dim x As RegExp',
    '    Set x = New RegExp',
    '    x.Pattern = "A"',
    '    Dim fso As Scripting.FileSystemObject',
    '    Set fso = New Scripting.FileSystemObject',
    '    Dim stream As Scripting.TextStream',
    '    If fso.FileExists("x") Then',
    '        Set stream = fso.GetFile("x").OpenAsTextStream(ForAppending)',
    '        stream.Close',
    '    End If',
    '    Debug.Print ForAppeding',
    'End Sub',
].join('\r\n');

suite('Scripting and RegExp project references', () => {
    let withRefs: string;
    let withoutRefs: string;
    suiteSetup(async () => {
        await activate();
        withRefs = path.join(workspaceRoot(), 'ScriptingWithReferences.xlsm');
        withoutRefs = path.join(workspaceRoot(), 'ScriptingWithoutReferences.xlsm');
        for (const file of [withRefs, withoutRefs]) {
            fs.copyFileSync(workbookPath(), file);
            writeModule(file, MODULE, SOURCE, 'standard');
        }
        const container = openMacroContainer(fs.readFileSync(withRefs));
        const cfb = container.vbaCfb();
        const project = VbaProject.parse(cfb);
        for (const ref of [
            { name: 'Scripting', guid: '{420B2830-E718-11CF-893D-00A0C9054228}', version: '1.0', path: 'C:\\Windows\\System32\\scrrun.dll', description: 'Microsoft Scripting Runtime' },
            { name: 'VBScript_RegExp_55', guid: '{3F4DACA7-160D-11D2-A8E9-00104B365C9F}', version: '5.5', path: 'C:\\Windows\\System32\\vbscript.dll\\3', description: 'Microsoft VBScript Regular Expressions 5.5' },
        ]) {
            project.addReferenceRecords(buildRegisteredReference(ref));
        }
        project.save(cfb);
        fs.writeFileSync(withRefs, container.toFileBytes(cfb));
        assert.ok(listReferences(withRefs).some(r => r.name === 'Scripting'));
        assert.ok(listReferences(withRefs).some(r => r.name === 'VBScript_RegExp_55'));
    });
    teardown(closeAllEditors);

    test('resolves constants and limits receiver completion with real workbook references', async () => {
        const doc = await open(encodeModuleUri(withRefs, MODULE));
        await until(() => vscode.languages.getDiagnostics(doc.uri)
            .find(d => d.code === 'undeclared-variable' && /ForAppeding/.test(d.message)),
        'the deliberate typo should establish that diagnostics have run');
        assert.ok(!vscode.languages.getDiagnostics(doc.uri).some(d => /ForAppending/.test(d.message)),
            'the referenced constant must not be reported undefined');
        for (const [marker, required, excluded] of [
            ['x.', 'Pattern', 'FileExists'], ['fso.', 'FileExists', 'Pattern'],
        ]) {
            const position = doc.positionAt(doc.getText().indexOf(marker) + marker.length);
            const names = await until(async () => {
                const list = await vscode.commands.executeCommand<vscode.CompletionList>(
                    'vscode.executeCompletionItemProvider', doc.uri, position, '.');
                const labels = list?.items.map(item => typeof item.label === 'string' ? item.label : item.label.label) ?? [];
                return labels.includes(required) ? labels : undefined;
            }, `completion should offer ${required}`);
            assert.ok(!names.includes(excluded));
            assert.ok(!names.includes('Workbook'));
        }
        const position = doc.positionAt(doc.getText().indexOf('ForAppending') + 3);
        const hovers = await vscode.commands.executeCommand<vscode.Hover[]>(
            'vscode.executeHoverProvider', doc.uri, position);
        const text = (hovers ?? []).flatMap(item => item.contents.map(content =>
            typeof content === 'string' ? content : content.value)).join('\n');
        assert.match(text, /ForAppending/);
        assert.match(text, /8/);
    });
    test('keeps the same constant undefined in a workbook without the reference', async () => {
        const doc = await open(encodeModuleUri(withoutRefs, MODULE));
        await until(() => vscode.languages.getDiagnostics(doc.uri)
            .find(d => d.code === 'undeclared-variable' && /ForAppending/.test(d.message)),
        'the constant should remain undefined without Microsoft Scripting Runtime');
    });
});