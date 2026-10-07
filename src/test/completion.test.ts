import * as assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { VbaEditorProjectContextService } from '../vbaEditorProjectContext';
import { VbaTypeSemanticTokensProvider } from '../vbaSemanticTokensProvider';
import type { VbaProjectIndexService } from '../vbaProjectIndexService';
import { backspaceNeedsExtension } from '../vbaEditorCommands';
import { activate, closeAllEditors, open, until, workspaceRoot, writeModule } from './support';

async function probe(name: string, source: string, marker: string): Promise<{ document: vscode.TextDocument; editor: vscode.TextEditor; caret: vscode.Position }> {
    const file = path.join(workspaceRoot(), `${name}.bas`);
    fs.writeFileSync(file, source);
    const document = await open(vscode.Uri.file(file));
    const editor = vscode.window.activeTextEditor!;
    const caret = document.positionAt(source.indexOf(marker) + marker.length);
    editor.selection = new vscode.Selection(caret, caret);
    await vscode.commands.executeCommand('workbench.action.focusActiveEditorGroup');
    return { document, editor, caret };
}
async function completions(document: vscode.TextDocument, caret: vscode.Position): Promise<vscode.CompletionList> {
    return (await vscode.commands.executeCommand<vscode.CompletionList>('vscode.executeCompletionItemProvider', document.uri, caret))!;
}
function labels(list: vscode.CompletionList): string[] {
    return list.items.map(item => typeof item.label === 'string' ? item.label : item.label.label);
}

suite('Completion editor surface', () => {
    const previousSettings = new Map<string, unknown>();
    suiteSetup(async () => {
        await activate();
        for (const key of ['quickSuggestions', 'wordBasedSuggestions', 'suggestSelection']) {
            previousSettings.set(key, vscode.workspace.getConfiguration('editor').inspect(key)?.workspaceValue);
        }
        await vscode.workspace.getConfiguration('editor').update('quickSuggestions', false, vscode.ConfigurationTarget.Workspace);
        await vscode.workspace.getConfiguration('editor').update('wordBasedSuggestions', 'off', vscode.ConfigurationTarget.Workspace);
        await vscode.workspace.getConfiguration('editor').update('suggestSelection', 'first', vscode.ConfigurationTarget.Workspace);
    });
    suiteTeardown(async () => {
        for (const [key, value] of previousSettings) {
            await vscode.workspace.getConfiguration('editor').update(key, value, vscode.ConfigurationTarget.Workspace);
        }
    });
    teardown(async () => {
        await vscode.commands.executeCommand('hideSuggestWidget');
        await vscode.commands.executeCommand('workbench.action.revertAndCloseActiveEditor');
        await closeAllEditors();
    });
    for (const trigger of ['=', ' ']) {
        test(`assignment constants popup automatically after ${JSON.stringify(trigger)} inside With`, async () => {
            const source = 'Sub Demo()\nDim ws As Worksheet\nWith ws.Range("A1")\n.HorizontalAlignment ' + (trigger === ' ' ? '=' : '') + '\nEnd With\nEnd Sub\n';
            const { document } = await probe(`AssignmentMenu${trigger === '=' ? 'Equals' : 'Space'}`, source, '.HorizontalAlignment ' + (trigger === ' ' ? '=' : ''));
            await vscode.commands.executeCommand('hideSuggestWidget');
            await vscode.commands.executeCommand('type', { text: trigger });
            await until(async () => {
                await vscode.commands.executeCommand('acceptSelectedSuggestion');
                return /\.HorizontalAlignment = ?xlHAlign/.test(document.lineAt(3).text) || undefined;
            }, 'typing the trigger must open the native enum menu even with quickSuggestions disabled', 4000);
        });
    }
    test('assignment constants retain variables and replace an existing value prefix once', async () => {
        const source = 'Sub Demo()\nDim chosen As XlHAlign\nActiveCell.HorizontalAlignment = xlHAlignLeft\nEnd Sub\n';
        const { document, caret } = await probe('AssignmentReplace', source, '= xlH');
        const list = await completions(document, caret);
        const item = list.items.find(item => item.label === 'xlHAlignCenter');
        assert.ok(item);
        assert.equal(list.items.filter(item => item.label === 'xlHAlignCenter').length, 1);
        const range = item.range instanceof vscode.Range ? item.range : item.range?.replacing;
        assert.ok(range);
        assert.equal(document.getText(range), 'xlHAlignLeft');
        const edit = new vscode.WorkspaceEdit();
        edit.replace(document.uri, range, 'xlHAlignCenter');
        await vscode.workspace.applyEdit(edit);
        assert.equal(document.lineAt(2).text, 'ActiveCell.HorizontalAlignment = xlHAlignCenter');
    });
    test('assignment constants stay responsive in a large module', async () => {
        const source = 'Sub Demo()\nDim value As Long\n' + 'value = value + 1\n'.repeat(3000) + 'ActiveCell.HorizontalAlignment = \nEnd Sub\n';
        const { document, caret } = await probe('AssignmentLarge', source, 'ActiveCell.HorizontalAlignment = ');
        const times: number[] = [];
        for (let i = 0; i < 11; i++) {
            const start = performance.now();
            assert.ok(labels(await completions(document, caret)).includes('xlHAlignLeft'));
            times.push(performance.now() - start);
        }
        const warm = times.slice(1).sort((a,b) => a-b);
        console.log(`Assignment completion provider command ms: first=${times[0].toFixed(1)}, warm median=${warm[5].toFixed(1)}, warm max=${warm.at(-1)!.toFixed(1)}`);
    });
    test('assignment diagnostics underline Height and clear after changing to RowHeight', async () => {
        const source = 'Option Explicit\nSub Demo(ByVal ws As Worksheet)\nws.Range("A1").EntireRow.Height = 20\nEnd Sub\n';
        const document = await open(await writeModule('AssignmentDiagnostics', source));
        const editor = vscode.window.activeTextEditor!;
        const diagnostic = await until(() => vscode.languages.getDiagnostics(document.uri).find(d=>d.code === 'host-readonly-value-assignment'), 'Height must receive an editor diagnostic', 5000);
        assert.equal(diagnostic.severity, vscode.DiagnosticSeverity.Error);
        assert.equal(document.getText(diagnostic.range), 'Height');
        assert.ok(diagnostic.message.includes('RowHeight'));
        await editor.edit(edit => edit.replace(diagnostic.range, 'RowHeight'));
        await until(() => vscode.languages.getDiagnostics(document.uri).every(d=>d.code !== 'host-readonly-value-assignment') || undefined, 'RowHeight must clear the read-only diagnostic', 5000);
    });
    test('assignment target bug hunt opens the native menu for an array element', async () => {
        const source = 'Option Explicit\nSub Demo()\nDim flags(1) As Boolean\nflags(1) \nEnd Sub\n';
        const { document } = await probe('ArrayAssignmentMenu', source, '\nflags(1) ');
        await vscode.commands.executeCommand('hideSuggestWidget');
        await vscode.commands.executeCommand('type', { text: '=' });
        await until(async () => {
            await vscode.commands.executeCommand('acceptSelectedSuggestion');
            return /flags\(1\) = ?(False|True)$/.test(document.lineAt(3).text) || undefined;
        }, `a Boolean array element must open its native value menu: ${document.lineAt(3).text}`, 4000);
    });
    test('assignment target bug hunt keeps a function-call equals quiet', async () => {
        const source = 'Option Explicit\nSub Demo()\nIsNumeric("abc") \nEnd Sub\n';
        const { document, editor } = await probe('FunctionAssignmentMenu', source, 'IsNumeric("abc") ');
        await vscode.commands.executeCommand('hideSuggestWidget');
        await vscode.commands.executeCommand('type', { text: '=' });
        const result = await vscode.commands.executeCommand<vscode.CompletionList>('vscode.executeCompletionItemProvider', document.uri, editor.selection.active, '=');
        assert.equal(result?.items.length, 0);
        await vscode.commands.executeCommand('acceptSelectedSuggestion');
        assert.equal(document.lineAt(2).text, 'IsNumeric("abc") =');
    });
    test('enum insertion bug hunt bypasses a shadowed library qualifier', async () => {
        const source = 'Sub Demo(ByVal sh As Shape)\nDim Office As Long\nDim msoTrue As Long\nsh.Visible = msoT\nEnd Sub\n';
        const { document, caret } = await probe('OfficeQualifierShadow', source, '= msoT');
        const item = (await completions(document, caret)).items.find(item => item.label === 'msoTrue');
        assert.ok(item);
        assert.equal(item.insertText, 'MsoTriState.msoTrue');
    });
    test('shared enum owner bug hunt qualifies a shadowed Office constant correctly', async () => {
        const source = 'Sub Demo(ByVal sh As Shape)\nDim msoTrue As Long\nsh.Visible = msoT\nEnd Sub\n';
        const { document, caret } = await probe('OfficeEnumOwner', source, '= msoT');
        const item = (await completions(document, caret)).items.find(item => item.label === 'msoTrue');
        assert.ok(item);
        assert.equal(item.insertText, 'Office.MsoTriState.msoTrue');
    });
    test('getter write bug hunt reports a scalar result and clears after an object return', async () => {
        const source = 'Option Explicit\nPublic Property Get GetterValue() As Variant\nGetterValue = 20\nEnd Property\nSub Demo()\nGetterValue = 30\nEnd Sub\n';
        const document = await open(await writeModule('GetterWriteDiagnostic', source));
        const editor = vscode.window.activeTextEditor!;
        const finding = await until(() => vscode.languages.getDiagnostics(document.uri).find(d => d.code === 'variant-value-misuse'), 'scalar getter result must report Object required', 5000);
        assert.equal(document.getText(finding.range), 'GetterValue');
        assert.equal(finding.severity, vscode.DiagnosticSeverity.Error);
        assert.ok(finding.message.includes("Run-time error '424'"));
        await editor.edit(edit => edit.replace(document.lineAt(2).range, 'Set GetterValue = ThisWorkbook.Worksheets(1).Range("A1")'));
        await until(() => !vscode.languages.getDiagnostics(document.uri).some(d => d.severity === vscode.DiagnosticSeverity.Error) || undefined, 'object return must clear the invalid getter write', 5000);
    });
    test('library shadow bug hunt keeps property diagnostics on the source receiver', async () => {
        const source = 'Option Explicit\nSub Demo(ByVal Word As Worksheet)\nWord.EnableCalculation = "nonsense"\nEnd Sub\n';
        const document = await open(await writeModule('LibraryShadowDiagnostic', source));
        const editor = vscode.window.activeTextEditor!;
        const finding = await until(() => vscode.languages.getDiagnostics(document.uri).find(d => d.code === 'assignment-type-mismatch'), 'Worksheet Boolean setter must reject invalid text', 5000);
        assert.equal(document.getText(finding.range), '"nonsense"');
        assert.equal(vscode.languages.getDiagnostics(document.uri).some(d => d.code === 'missing-library-reference'), false);
        await editor.edit(edit => edit.replace(finding.range, 'True'));
        await until(() => !vscode.languages.getDiagnostics(document.uri).some(d => d.severity === vscode.DiagnosticSeverity.Error) || undefined, 'valid source-bound property write must clear errors', 5000);
    });
    test('bare setter bug hunt reports an invalid value and clears on correction', async () => {
        const source = 'Option Explicit\nPublic Property Let State(ByVal value As Boolean)\nEnd Property\nSub Demo()\nState = "nonsense"\nEnd Sub\n';
        const document = await open(await writeModule('BareSetterDiagnostic', source));
        const editor = vscode.window.activeTextEditor!;
        const finding = await until(() => vscode.languages.getDiagnostics(document.uri).find(d=>d.code === 'assignment-type-mismatch'), 'bare setter must report a type mismatch', 5000);
        assert.equal(document.getText(finding.range), '"nonsense"');
        assert.equal(finding.severity, vscode.DiagnosticSeverity.Error);
        await editor.edit(edit => edit.replace(finding.range, 'True'));
        await until(() => !vscode.languages.getDiagnostics(document.uri).some(d=>d.code === 'assignment-type-mismatch') || undefined, 'valid Boolean assignment must clear the finding', 5000);
    });
    test('enum coercion bug hunt reports invalid text and accepts an unnamed number', async () => {
        const source = 'Option Explicit\nSub Demo()\nDim alignment As XlHAlign\nalignment = "abc"\nDebug.Print alignment\nEnd Sub\n';
        const document = await open(await writeModule('EnumCoercionDiagnostic', source));
        const editor = vscode.window.activeTextEditor!;
        const finding = await until(() => vscode.languages.getDiagnostics(document.uri).find(d=>d.code === 'assignment-type-mismatch'), 'host enum must reject nonnumeric text', 5000);
        assert.equal(document.getText(finding.range), '"abc"');
        await editor.edit(edit => edit.replace(finding.range, '999'));
        await until(() => !vscode.languages.getDiagnostics(document.uri).some(d=>d.code === 'assignment-type-mismatch') || undefined, 'unnamed numeric enum value must remain valid', 5000);
    });
    test('exported getter bug hunt opens its field value menu and clears a mismatch', async () => {
        await open(await writeModule('ExportGetterTypes', 'Public Type GetterRecord\nFlag As Boolean\nEnd Type\nPublic Property Get ExportedSnapshot() As GetterRecord\nEnd Property\n'));
        const document = await open(await writeModule('ExportGetterCaller', 'Option Explicit\nSub Demo()\nExportedSnapshot.Flag = "nonsense"\nEnd Sub\n'));
        const editor = vscode.window.activeTextEditor!;
        const finding = await until(() => vscode.languages.getDiagnostics(document.uri).find(d=>d.code === 'assignment-type-mismatch'), 'bare exported getter field must reject invalid text', 5000);
        assert.equal(document.getText(finding.range), '"nonsense"');
        await editor.edit(edit => edit.replace(finding.range, 'True'));
        await until(() => !vscode.languages.getDiagnostics(document.uri).some(d=>d.severity === vscode.DiagnosticSeverity.Error) || undefined, 'valid returned-field assignment must clear errors', 5000);
        const times: number[] = [];
        for (let repeat = 0; repeat < 5; repeat++) {
            const assignment = document.lineAt(2).text;
            const equals = assignment.indexOf('=');
            await editor.edit(edit => edit.replace(new vscode.Range(2, equals, 2, assignment.length), ''));
            editor.selection = new vscode.Selection(2, equals, 2, equals);
            await vscode.commands.executeCommand('workbench.action.focusActiveEditorGroup');
            await vscode.commands.executeCommand('hideSuggestWidget');
            const start = performance.now();
            await vscode.commands.executeCommand('type', { text: '=' });
            times.push(performance.now() - start);
            await until(async () => {
                await vscode.commands.executeCommand('acceptSelectedSuggestion');
                return /Flag = ?(False|True)$/.test(document.lineAt(2).text) || undefined;
            }, 'exported getter field must open its native Boolean menu', 4000);
        }
        console.log(`Exported getter equals command ms (loaded project, first then repeated edits): [${times.map(time => time.toFixed(1)).join(',')}]`);
    });
    test('indexed Set bug hunt reports an object mismatch and clears on correction', async () => {
        const source = 'Option Explicit\nProperty Set Item(ByVal index As Long, ByVal value As Worksheet)\nEnd Property\nSub Demo(ByVal ws As Worksheet)\nSet IndexedSetDiagnostic.Item(1) = New Collection\nEnd Sub\n';
        const document = await open(await writeModule('IndexedSetDiagnostic', source));
        const editor = vscode.window.activeTextEditor!;
        const finding = await until(() => vscode.languages.getDiagnostics(document.uri).find(d=>d.code === 'assignment-object-type-mismatch'), 'indexed setter must report incompatible objects', 5000);
        assert.equal(finding.severity, vscode.DiagnosticSeverity.Error);
        const start = document.getText().indexOf('New Collection');
        await editor.edit(edit => edit.replace(new vscode.Range(document.positionAt(start), document.positionAt(start + 'New Collection'.length)), 'ws'));
        await until(() => !vscode.languages.getDiagnostics(document.uri).some(d=>d.severity === vscode.DiagnosticSeverity.Error) || undefined, 'valid Worksheet must clear the indexed setter finding', 5000);
    });
    test('array setter bug hunt clears the comparison error for a valid With Range.Value transfer', async () => {
        const source = 'Option Explicit\nSub Demo()\nDim values As Variant\nvalues = Array(1,2)\nIf values = 1 Then Exit Sub\nEnd Sub\n';
        const document = await open(await writeModule('WithArrayTransferDiagnostic', source));
        const editor = vscode.window.activeTextEditor!;
        await until(() => vscode.languages.getDiagnostics(document.uri).find(d=>d.code === 'variant-value-misuse'), 'a real array comparison must be diagnosed', 5000);
        await editor.edit(edit => edit.replace(new vscode.Range(4, 0, 4, document.lineAt(4).text.length), 'With ThisWorkbook.Worksheets(1).Range("A1:B1")\n.Value = values\nEnd With'));
        await until(() => !vscode.languages.getDiagnostics(document.uri).some(d=>d.severity === vscode.DiagnosticSeverity.Error) || undefined, 'valid With array transfer must clear the comparison diagnostic', 5000);
    });
    test('array setter bug hunt keeps Byte-array runtime errors separate from valid String conversion', async () => {
        const source = 'Option Explicit\nProperty Let Payload(ByRef bytes() As Byte)\nEnd Property\nSub Demo()\nDim value As Variant\nvalue = Array(1,2)\nPayload = value\nEnd Sub\n';
        const document = await open(await writeModule('ByteArraySetterDiagnostic', source));
        const editor = vscode.window.activeTextEditor!;
        const finding = await until(() => vscode.languages.getDiagnostics(document.uri).find(d=>d.code === 'assignment-type-mismatch'), 'tracked Variant array must report runtime type mismatch', 5000);
        assert.ok(finding.message.includes("Run-time error '13'"));
        assert.equal(vscode.languages.getDiagnostics(document.uri).some(d=>d.code === 'argument-shape-mismatch'), false);
        const start = document.getText().indexOf('Array(1,2)');
        await editor.edit(edit => edit.replace(new vscode.Range(document.positionAt(start), document.positionAt(start + 'Array(1,2)'.length)), '"text"'));
        await until(() => !vscode.languages.getDiagnostics(document.uri).some(d=>d.severity === vscode.DiagnosticSeverity.Error) || undefined, 'Variant String conversion must clear the Byte-array finding', 5000);
    });
    test('array setter bug hunt reports a compile shape error and accepts a typed function result', async () => {
        const source = 'Option Explicit\nProperty Let Flags(ByRef value() As Boolean)\nEnd Property\nFunction MakeFlags() As Boolean()\nDim result(1) As Boolean\nMakeFlags = result\nEnd Function\nSub Demo()\nFlags = True\nEnd Sub\n';
        const document = await open(await writeModule('ArraySetterDiagnostic', source));
        const editor = vscode.window.activeTextEditor!;
        const finding = await until(() => vscode.languages.getDiagnostics(document.uri).find(d=>d.code === 'argument-shape-mismatch'), 'array setter must reject the scalar at compile time', 5000);
        assert.equal(document.getText(finding.range), 'True');
        assert.ok(finding.message.includes('VBE compile error'));
        assert.equal(vscode.languages.getDiagnostics(document.uri).some(d=>d.code === 'assignment-type-mismatch'), false);
        await editor.edit(edit => edit.replace(finding.range, 'MakeFlags()'));
        await until(() => !vscode.languages.getDiagnostics(document.uri).some(d=>d.severity === vscode.DiagnosticSeverity.Error) || undefined, 'typed array function must clear the shape finding', 5000);
    });
    test('setter shape bug hunt opens a DefBool setter menu automatically', async () => {
        const source = 'DefBool V\nProperty Let State(ByVal value)\nEnd Property\nSub Demo()\nState \nEnd Sub\n';
        const { document } = await probe('DefBoolSetterMenu', source, '\nState ');
        await vscode.commands.executeCommand('hideSuggestWidget');
        await vscode.commands.executeCommand('type', { text: '=' });
        await until(async () => {
            await vscode.commands.executeCommand('acceptSelectedSuggestion');
            return /State = ?(False|True)$/.test(document.lineAt(4).text) || undefined;
        }, 'the implicitly Boolean setter must open its native value menu', 4000);
    });
    test('scalar setter array bug hunt reports a whole array and accepts its element', async () => {
        const source = 'Option Explicit\nPublic Property Let State(ByVal value As Boolean)\nEnd Property\nSub Demo()\nDim values(1) As Boolean\nState = values\nEnd Sub\n';
        const document = await open(await writeModule('ScalarSetterArrayDiagnostic', source));
        const editor = vscode.window.activeTextEditor!;
        const finding = await until(() => vscode.languages.getDiagnostics(document.uri).find(d => d.code === 'array-assignment-to-scalar'), 'scalar setter must reject a whole typed array', 5000);
        assert.equal(document.getText(finding.range), 'values');
        assert.ok(finding.message.includes('VBE compile error'));
        await editor.edit(edit => edit.replace(finding.range, 'values(0)'));
        await until(() => !vscode.languages.getDiagnostics(document.uri).some(d => d.severity === vscode.DiagnosticSeverity.Error) || undefined, 'array element must clear the scalar setter error', 5000);
    });
    test('vb type name bug hunt retains the source enum Long range', async () => {
        const source = 'Enum VbInteger\nFirst = 1\nEnd Enum\nProperty Let State(ByVal value As VbInteger)\nEnd Property\nSub Demo()\nState = 2147483648#\nEnd Sub\n';
        const document = await open(await writeModule('VbNamedEnumDiagnostic', source));
        const editor = vscode.window.activeTextEditor!;
        await until(() => vscode.languages.getDiagnostics(document.uri).find(d => d.code === 'assignment-type-mismatch'), 'out-of-Long-range enum value must fail', 5000);
        await editor.edit(edit => edit.replace(document.lineAt(6).range, 'State = 50000'));
        await until(() => !vscode.languages.getDiagnostics(document.uri).some(d => d.severity === vscode.DiagnosticSeverity.Error) || undefined, 'source enum must allow a Long value beyond Integer range', 5000);
    });
    test('qualified array bug hunt refreshes array element compatibility', async () => {
        const source = 'Function Factory() As Long()\nDim data(1) As Long\nFactory = data\nEnd Function\nSub Demo()\nDim values() As Boolean\nvalues = QualifiedArrayDiagnostic.Factory()\nEnd Sub\n';
        const document = await open(await writeModule('QualifiedArrayDiagnostic', source));
        const editor = vscode.window.activeTextEditor!;
        const finding = await until(() => vscode.languages.getDiagnostics(document.uri).find(d => d.code === 'array-target-assignment'), 'qualified Long array must not fit a Boolean array', 5000);
        assert.equal(document.getText(finding.range), 'values');
        await editor.edit(edit => edit.replace(document.lineAt(5).range, 'Dim values() As Long'));
        await until(() => !vscode.languages.getDiagnostics(document.uri).some(d => d.severity === vscode.DiagnosticSeverity.Error) || undefined, 'matching array element types must clear the error', 5000);
    });
    test('function default bug hunt refreshes the returned default contract', async () => {
        const source = 'Function Factory() As Collection\nSet Factory = New Collection\nEnd Function\nSub Demo()\nFactory() = 20\nEnd Sub\n';
        const document = await open(await writeModule('FactoryDefaultDiagnostic', source));
        const editor = vscode.window.activeTextEditor!;
        const finding = await until(() => vscode.languages.getDiagnostics(document.uri).find(d => d.code === 'argument-count'), 'Collection result must require a default index', 5000);
        assert.equal(document.getText(finding.range), 'Factory');
        const corrected = source.replace('As Collection', 'As Range').replace('New Collection', 'ThisWorkbook.Worksheets(1).Range("A1")');
        await editor.edit(edit => edit.replace(new vscode.Range(document.positionAt(0), document.positionAt(document.getText().length)), corrected));
        await until(() => !vscode.languages.getDiagnostics(document.uri).some(d => d.severity === vscode.DiagnosticSeverity.Error) || undefined, 'Range result must remain writable', 5000);
    });
    test('getter default bug hunt checks and refreshes the returned default contract', async () => {
        const source = 'Option Explicit\nPublic Property Get Child() As Collection\nSet Child = New Collection\nEnd Property\nSub Demo()\nChild = 20\nEnd Sub\n';
        const document = await open(await writeModule('GetterDefaultDiagnostic', source));
        const editor = vscode.window.activeTextEditor!;
        const finding = await until(() => vscode.languages.getDiagnostics(document.uri).find(d => d.code === 'argument-count'), 'returned Collection must require its default index', 5000);
        assert.equal(document.getText(finding.range), 'Child');
        await editor.edit(edit => {
            edit.replace(document.lineAt(1).range, 'Public Property Get Child() As Range');
            edit.replace(document.lineAt(2).range, 'Set Child = ThisWorkbook.Worksheets(1).Range("A1")');
        });
        await until(() => !vscode.languages.getDiagnostics(document.uri).some(d => d.severity === vscode.DiagnosticSeverity.Error) || undefined, 'returned Range must remain writable through its default', 5000);
    });
    test('receiver contract bug hunt requires ByRef exactness in a With header', async () => {
        const source = 'Option Explicit\nPublic Function GetSheet(ByRef index As Long) As Worksheet\nSet GetSheet = ThisWorkbook.Worksheets(1)\nEnd Function\nSub Demo()\nDim i As Integer\nWith GetSheet(i)\n.EnableCalculation = True\nEnd With\nEnd Sub\n';
        const document = await open(await writeModule('ReceiverContractDiagnostic', source));
        const editor = vscode.window.activeTextEditor!;
        const finding = await until(() => vscode.languages.getDiagnostics(document.uri).find(d => d.code === 'byref-argument-type-mismatch'), 'With receiver must require an exact ByRef type', 5000);
        assert.equal(document.getText(finding.range), 'i');
        await editor.edit(edit => edit.replace(finding.range, '(i)'));
        await until(() => !vscode.languages.getDiagnostics(document.uri).some(d => d.severity === vscode.DiagnosticSeverity.Error) || undefined, 'copy conversion must clear the With receiver diagnostic', 5000);
    });
    test('setter index type bug hunt refreshes coercion and ByRef diagnostics', async () => {
        const source = 'Option Explicit\nPublic Property Let State(ByVal index As Long, ByVal value As Boolean)\nEnd Property\nSub Demo()\nDim i As Integer\nState("bad") = True\nEnd Sub\n';
        const document = await open(await writeModule('SetterIndexTypeDiagnostic', source));
        const editor = vscode.window.activeTextEditor!;
        const literal = await until(() => vscode.languages.getDiagnostics(document.uri).find(d => d.code === 'argument-type-mismatch'), 'setter index must reject nonnumeric text', 5000);
        assert.equal(document.getText(literal.range), '"bad"');
        await editor.edit(edit => {
            edit.replace(document.lineAt(1).range, 'Public Property Let State(ByRef index As Long, ByVal value As Boolean)');
            edit.replace(literal.range, 'i');
        });
        const mismatch = await until(() => vscode.languages.getDiagnostics(document.uri).find(d => d.code === 'byref-argument-type-mismatch'), 'changed setter passing mode must require an exact variable type', 5000);
        assert.equal(document.getText(mismatch.range), 'i');
        await editor.edit(edit => edit.replace(mismatch.range, '(i)'));
        await until(() => !vscode.languages.getDiagnostics(document.uri).some(d => d.severity === vscode.DiagnosticSeverity.Error) || undefined, 'parenthesized conversion must clear the index error', 5000);
    });
    test('setter index bug hunt reports a missing index and clears on correction', async () => {
        const source = 'Option Explicit\nPublic Property Let State(ByVal index As Long, ByVal value As Boolean)\nEnd Property\nSub Demo()\nState = True\nEnd Sub\n';
        const document = await open(await writeModule('SetterIndexDiagnostic', source));
        const editor = vscode.window.activeTextEditor!;
        const finding = await until(() => vscode.languages.getDiagnostics(document.uri).find(d => d.code === 'argument-count'), 'setter-only index must be required', 5000);
        assert.equal(document.getText(finding.range), 'State');
        assert.ok(finding.message.includes('Argument not optional'));
        await editor.edit(edit => edit.insert(finding.range.end, '(1)'));
        await until(() => !vscode.languages.getDiagnostics(document.uri).some(d => d.severity === vscode.DiagnosticSeverity.Error) || undefined, 'valid setter index must clear the error', 5000);
    });
    test('setter shape bug hunt keeps scalar constants out of an array value slot', async () => {
        const source = 'Property Let Flags(ByRef value() As Boolean)\nEnd Property\nSub Demo()\nFlags \nEnd Sub\n';
        const { document, editor } = await probe('ArraySetterMenu', source, '\nFlags ');
        await vscode.commands.executeCommand('hideSuggestWidget');
        await vscode.commands.executeCommand('type', { text: '=' });
        const result = await vscode.commands.executeCommand<vscode.CompletionList>('vscode.executeCompletionItemProvider', document.uri, editor.selection.active, '=');
        assert.equal(result?.items.length, 0);
    });
    test('invalidates semantic tokens across real document close/open language events', async () => {
        const provider = new VbaTypeSemanticTokensProvider({} as VbaProjectIndexService);
        const token = { isCancellationRequested: false } as vscode.CancellationToken;
        try {
            const { document } = await probe('SemanticReopen',
                'Sub Demo()\n    Debug.Print ThisWorkbook.Name\nEnd Sub\n', 'ThisWorkbook');
            const before = await provider.provideDocumentSemanticTokens(document, token);
            assert.ok(before.data.length, 'initial file should have semantic tokens');
            const version = document.version;
            const language = document.languageId;
            // Language changes emit close/open document events even when the
            // workbench retains the file model after a tab closes.
            const plain = await vscode.languages.setTextDocumentLanguage(document, 'plaintext');
            const reopened = await vscode.languages.setTextDocumentLanguage(plain, language);
            assert.equal(reopened.version, version, 'document lifecycle should preserve the same cache version');
            const after = await provider.provideDocumentSemanticTokens(reopened, token);
            assert.ok(after.data.length, 'reopened VBA document should have semantic tokens');
            assert.notEqual(after, before, 'closed document token cache should have been discarded');
            assert.deepEqual(Array.from(after.data), Array.from(before.data), 'unchanged source should retain its token positions');
        } finally { provider.dispose(); }
    });
    test('does not reuse completion contexts across real document close/open language events', async () => {
        const service = new VbaEditorProjectContextService({} as VbaProjectIndexService);
        const { document } = await probe('ContextReopen',
            'Public Type ProbeType\nValue As Long\nEnd Type\n', 'ProbeType');
        const before = await service.buildEditorProjectContext(document, document.getText());
        assert.ok(before.projectTypes?.some(type => type.name === 'ProbeType'));
        const version = document.version;
        const language = document.languageId;
        const plain = await vscode.languages.setTextDocumentLanguage(document, 'plaintext');
        const reopened = await vscode.languages.setTextDocumentLanguage(plain, language);
        assert.equal(reopened.version, version);
        assert.equal(service.cachedEditorProjectContext(reopened), undefined);
        const after = await service.buildEditorProjectContext(reopened, reopened.getText());
        assert.notEqual(after, before);
        assert.ok(after.projectTypes?.some(type => type.name === 'ProbeType'));
        service.dispose();
    });
    test('Smart Enter seeds a With member and undo restores the original source', async () => {
        const source = 'Sub Demo()\n\tWith ActiveSheet\n\tEnd With\nEnd Sub\n';
        const { document, editor } = await probe('SmartEnterUndo', source, 'With ActiveSheet');
        await vscode.commands.executeCommand('type', { text: '\n' });
        await until(() => {
            const caret = editor.selection.active;
            return document.lineAt(caret.line).text.slice(0, caret.character).endsWith('.') || undefined;
        }, 'Smart Enter should place the caret after the seeded With dot', 4000);
        await vscode.commands.executeCommand('hideSuggestWidget');
        await vscode.commands.executeCommand('undo');
        try {
            await until(() => document.getText() === source || undefined,
                'undo should restore the source without generating another Smart Enter edit', 4000);
        } catch (error) {
            throw new Error(`${String(error)}; actual=${JSON.stringify(document.getText())}`);
        }
    });
    test('Smart Backspace clears a continued comment then its remaining indent', async () => {
        const { document, editor } = await probe('CommentBackspace',
            "Sub Demo()\n    'note\n    ' \nEnd Sub\n", "    ' ");
        assert.equal(backspaceNeedsExtension(editor), true);
        await vscode.commands.executeCommand('xlide.vba.smartBackspace');
        assert.equal(document.lineAt(2).text, '    ');
        await vscode.commands.executeCommand('xlide.vba.smartBackspace');
        assert.equal(document.lineAt(2).text, '');
    });
    test('Smart Tab clears a continued comment and indents the blank line', async () => {
        const { document } = await probe('CommentTab',
            "Sub Demo()\n    'note\n    ' \nEnd Sub\n", "    ' ");
        await vscode.commands.executeCommand('xlide.vba.smartTab');
        assert.equal(document.lineAt(2).text.trim(), '');
        assert.ok(document.lineAt(2).text.length > 4, 'Tab should deepen the remaining indent');
    });
    test('serves repeated completion requests in a large unchanged module', async () => {
        const source = 'Sub Demo()\nDim value As Long\n' + 'value = value + 1\n'.repeat(3000) +
            'ThisWorkbook.Sheets(1).ce\nEnd Sub\n';
        const { document, caret } = await probe('CompletionLarge', source, '.ce');
        const times: number[] = [];
        for (let i = 0; i < 10; i++) {
            const start = performance.now();
            assert.ok(labels(await completions(document, caret)).includes('Cells'));
            times.push(performance.now() - start);
        }
        console.log(`Completion large module ms: first=${times[0].toFixed(1)}, repeated=[${times.slice(1).map(time => time.toFixed(1)).join(',')}]`);
    });
    test('returns Cells for a corrected worksheet prefix', async () => {
        const { document, editor, caret } = await probe('CompletionPrefix', 'Sub Demo()\nThisWorkbook.Sheets(1).cez\nEnd Sub\n', '.cez');
        assert.deepEqual(labels(await completions(document, caret)), []);
        const start = caret.translate(0, -1);
        await editor.edit(edit => edit.delete(new vscode.Range(start, caret)));
        editor.selection = new vscode.Selection(start, start);
        assert.ok(labels(await completions(document, start)).includes('Cells'));
    });
    test('reopens the actual menu after Backspace, allowing Cells to be accepted', async () => {
        const { document } = await probe('CompletionMenu', 'Sub Demo()\nThisWorkbook.Sheets(1).cez\nEnd Sub\n', '.cez');
        await vscode.commands.executeCommand('hideSuggestWidget');
        await vscode.commands.executeCommand('deleteLeft');
        await until(async () => {
            await vscode.commands.executeCommand('acceptSelectedSuggestion');
            return document.lineAt(1).text.endsWith('.Cells') ? true : undefined;
        },
            `backspace should reopen Cells: ${document.lineAt(1).text}`, 4000);
    });
    test('offers procedures from the current module inside a macro-name string', async () => {
        const source = 'Public Sub Clicked()\nEnd Sub\nSub Demo()\nApplication.Run ""\nEnd Sub\n';
        const { document, caret } = await probe('CompletionMacro', source, 'Application.Run "');
        assert.ok(labels(await completions(document, caret)).includes('Module.Clicked'),
            'Application.Run should offer the current module procedure');
    });
    test('preserves both quotes when a macro suggestion replaces an empty string', async () => {
        const source = 'Public Sub Clicked()\nEnd Sub\nSub Demo()\nApplication.Run ""\nEnd Sub\n';
        const { document, caret } = await probe('CompletionMacroQuotes', source, 'Application.Run "');
        const item = (await completions(document, caret)).items.find(candidate => candidate.label === 'Module.Clicked');
        assert.ok(item);
        const range = item.range instanceof vscode.Range ? item.range : item.range?.replacing;
        assert.ok(range);
        assert.equal(document.getText(range), '', 'the completion range must exclude the closing quote');
        const edit = new vscode.WorkspaceEdit();
        edit.replace(document.uri, range, 'Module.Clicked');
        await vscode.workspace.applyEdit(edit);
        assert.equal(document.lineAt(3).text, 'Application.Run "Module.Clicked"');
    });
    test('hovers a current-module procedure named in a macro string', async () => {
        const source = 'Public Sub Clicked()\nEnd Sub\nSub Demo()\nApplication.Run "Module.Clicked"\nEnd Sub\n';
        const { document, caret } = await probe('CompletionMacroHover', source, '"Module.Clicked');
        const hovers = await vscode.commands.executeCommand<vscode.Hover[]>('vscode.executeHoverProvider', document.uri, caret.translate(0, -1));
        assert.ok(hovers?.some(hover => hover.contents.some(content =>
            (typeof content === 'string' ? content : content.value).includes('Sub Clicked'))),
            'hover should resolve the current module procedure named by Application.Run');
    });
    for (const [expression, name, prefix] of [
        ['value = Abs(-1)', 'Abs', 'value = Ab'],
        ['value = Left$("abc", 1)', 'Left$', 'value = Lef'],
        ['Set value = Application.Intersect(a, b)', 'Intersect', 'Application.Int'],
    ]) {
        test(`preserves existing arguments when completing ${name}`, async () => {
            const source = `Sub Demo()\n${expression}\nEnd Sub\n`;
            const { document, editor, caret } = await probe(`CompletionArguments${name}`, source, prefix);
            const item = (await completions(document, caret)).items.find(candidate => candidate.label === name);
            assert.ok(item);
            const range = item.range instanceof vscode.Range ? item.range : item.range?.replacing;
            assert.ok(range);
            const insertText = item.insertText;
            if (insertText instanceof vscode.SnippetString) {
                await editor.insertSnippet(insertText, range);
            } else {
                await editor.edit(edit => edit.replace(range, insertText ?? name));
            }
            assert.equal(document.lineAt(1).text, expression, 'completion must retain the original argument list');
        });
    }
    test('completes and accepts a bracketed worksheet member in the actual menu', async () => {
        const source = 'Sub Demo()\nThisWorkbook.Sheets(1).[Ce]\nEnd Sub\n';
        const { document, caret } = await probe('CompletionBracketed', source, '.[Ce');
        const item = (await completions(document, caret)).items.find(candidate => candidate.label === 'Cells');
        assert.ok(item, 'Cells must be offered inside a bracketed name');
        const range = item.range instanceof vscode.Range ? item.range : item.range?.replacing;
        assert.ok(range);
        assert.equal(document.getText(range), '[Ce]');
        await vscode.commands.executeCommand('editor.action.triggerSuggest');
        await until(async () => {
            await vscode.commands.executeCommand('acceptSelectedSuggestion');
            return document.lineAt(1).text.endsWith('.[Cells]') ? true : undefined;
        },
            'accepting Cells must replace the bracketed name once', 4000);
    });
    test('reopens the bracketed member menu after an unmatched prefix is corrected', async () => {
        const source = 'Sub Demo()\nThisWorkbook.Sheets(1).[Cez]\nEnd Sub\n';
        const { document } = await probe('CompletionBracketedRecovery', source, '.[Cez');
        await vscode.commands.executeCommand('hideSuggestWidget');
        await vscode.commands.executeCommand('deleteLeft');
        await until(async () => {
            await vscode.commands.executeCommand('acceptSelectedSuggestion');
            return document.lineAt(1).text.endsWith('.[Cells]') ? true : undefined;
        }, 'Backspace must reopen the bracketed member menu', 4000);
    });
    test('inserts Err as an object rather than an empty function call', async () => {
        const { document, editor, caret } = await probe('CompletionRuntimeObject', 'Sub Demo()\nSet obj = Er\nEnd Sub\n', 'Set obj = Er');
        const item = (await completions(document, caret)).items.find(candidate => candidate.label === 'Err');
        assert.ok(item);
        const range = item.range instanceof vscode.Range ? item.range : item.range?.replacing;
        assert.ok(range);
        const insertText = item.insertText;
        if (insertText instanceof vscode.SnippetString) {
            await editor.insertSnippet(insertText, range);
        } else {
            await editor.edit(edit => edit.replace(range, insertText ?? 'Err'));
        }
        assert.equal(document.lineAt(1).text, 'Set obj = Err');
    });
    test('inserts parentheses for a bare host method in an expression', async () => {
        const { document, editor, caret } = await probe('CompletionGlobalMethod', 'Sub Demo()\nSet obj = Uni\nEnd Sub\n', 'Set obj = Uni');
        const item = (await completions(document, caret)).items.find(candidate => candidate.label === 'Union');
        assert.ok(item);
        assert.ok(item.insertText instanceof vscode.SnippetString);
        const range = item.range instanceof vscode.Range ? item.range : item.range?.replacing;
        assert.ok(range);
        await editor.insertSnippet(item.insertText, range);
        assert.equal(document.lineAt(1).text, 'Set obj = Union()');
    });
    test('keeps ordinary member-prefix deletion on the native Backspace route', async () => {
        const expression = 'ThisWorkbook.Sheets(1).az';
        const source = 'Public Property Get Demo() As Variant\nIf True Then\nEnd If\n' + expression + '\nEnd Property\n';
        const { document, editor } = await probe('CompletionRepeatedBackspace', source, expression);
        for (let removed = 1; removed <= 10; removed++) {
            assert.equal(backspaceNeedsExtension(editor), false);
            await vscode.commands.executeCommand('deleteLeft');
            assert.equal(document.lineAt(3).text, expression.slice(0, -removed), `Backspace ${removed} must delete another character`);
        }
    });
    test('keeps smart Backspace working while joining trailing blank lines', async () => {
        const source = 'Sub Demo()\nEnd Sub\n' + '\n'.repeat(12);
        const { document } = await probe('CompletionBackspaceLineJoin', source, 'End Sub');
        await vscode.commands.executeCommand('cursorBottom');
        for (let removed = 1; removed <= 12; removed++) {
            await vscode.commands.executeCommand('xlide.vba.smartBackspace');
            assert.equal(document.getText(), source.slice(0, -removed), `Backspace ${removed} must join the next blank line`);
        }
    });
    test('measures typing and smart Backspace in a large module', async () => {
        const source = 'Sub Demo()\nDim value As Long\n' + 'value = value + 1\n'.repeat(3000) + 'value = 12345\nEnd Sub\n';
        const { document } = await probe('CompletionTypingWork', source, 'value = 12345');
        const times: number[] = [];
        for (const text of ['6', '7', '8', '9', '0']) {
            const start = performance.now();
            await vscode.commands.executeCommand('type', { text });
            times.push(performance.now() - start);
        }
        for (let removed = 0; removed < 5; removed++) {
            await vscode.commands.executeCommand('xlide.vba.smartBackspace');
        }
        assert.equal(document.lineAt(3002).text, 'value = 12345');
        console.log(`Typing large module ms: [${times.map(time => time.toFixed(1)).join(',')}]`);
    });
    test('does not offer code completions inside an ordinary string', async () => {
        const { document, caret } = await probe('CompletionString', 'Sub Demo()\nDebug.Print "hello"\nEnd Sub\n', 'hello');
        assert.equal((await completions(document, caret)).items.length, 0);
    });
    test('does not offer code completions in a comment', async () => {
        const { document, caret } = await probe('CompletionComment', "Sub Demo()\n' ordinary comment\nEnd Sub\n", 'comment');
        assert.equal((await completions(document, caret)).items.length, 0);
    });
    test('inserts a member containing combining marks as an ordinary identifier', async () => {
        const name = '\u0915\u093eValue';
        const source = `Public Type Record\n${name} As Long\nEnd Type\nSub Demo()\nDim obj As Record\nobj.\u0915\u093e\nEnd Sub\n`;
        const { document, caret } = await probe('CompletionCombiningMarks', source, 'obj.\u0915\u093e');
        const item = (await completions(document, caret)).items.find(candidate => candidate.label === name);
        assert.ok(item);
        assert.equal(item.insertText, name, 'a valid combining-mark identifier must not be bracketed');
    });
    test('replaces a Unicode UDT member prefix with the correct editor range', async () => {
        const source = 'Public Type Record\nCaf\u00e9Value As Long\nEnd Type\nSub Demo()\nDim obj As Record\nobj.caf\u00e9\nEnd Sub\n';
        const { document, caret } = await probe('CompletionUnicode', source, 'obj.caf\u00e9');
        const list = await completions(document, caret);
        const item = list.items.find(item => item.label === 'Caf\u00e9Value');
        assert.ok(item, 'UDT field should be offered');
        const range = item.range instanceof vscode.Range ? item.range : item.range?.replacing;
        assert.ok(range);
        assert.equal(document.getText(range), 'caf\u00e9');
    });
});
