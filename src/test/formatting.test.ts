// Format Document in a real VS Code: the provider is registered for both
// faces of VBA (a workbook module and a loose file), it answers through the
// same command the editor calls, and it leaves a designer/attribute header
// alone. Documents are opened without editors: the text read here must be
// the text the edits were computed for, and an editor invites the casing
// controller to change a line in between.

import * as assert from 'node:assert/strict';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { readModules } from '../vba/projectService';
import { activate, applyEdits, closeAllEditors, openHidden, until, workbookPath, workspaceRoot, writeModule } from './support';

const OPTIONS: vscode.FormattingOptions = { tabSize: 4, insertSpaces: true };

async function formatEdits(document: vscode.TextDocument): Promise<vscode.TextEdit[]> {
	const edits = await vscode.commands.executeCommand<vscode.TextEdit[]>(
		'vscode.executeFormatDocumentProvider', document.uri, OPTIONS,
	);
	return edits ?? [];
}

suite('Format Document', () => {
	suiteSetup(async () => {
		await activate();
	});

	teardown(async () => {
		await closeAllEditors();
	});

	test('re-indents, cases keywords and spaces a workbook module', async () => {
		const uri = await writeModule('FormatProbe', [
			'option explicit',
			'public sub greet(byval who as string)',
			'dim msg as string',
			'if len(who)>0 then',
			'msg="Hello, "&who',
			'debug.print msg',
			'end if',
			'end sub',
			'',
		].join('\r\n'));
		const document = await openHidden(uri);
		const edits = await formatEdits(document);
		assert.ok(edits.length > 0, 'the formatter should have something to change');
		assert.equal(applyEdits(document.getText(), edits), [
			'Option Explicit',
			'Public Sub greet(ByVal who As String)',
			'    Dim msg As String',
			'    If Len(who) > 0 Then',
			'        msg = "Hello, " & who',
			'        Debug.Print msg',
			'    End If',
			'End Sub',
			'',
		].join('\r\n'));
	});

	test('answers with no edits for a module that is already formatted', async () => {
		const uri = await writeModule('FormatClean', 'Option Explicit\r\n\r\nSub T()\r\n    Debug.Print 1\r\nEnd Sub\r\n');
		const document = await openHidden(uri);
		assert.deepEqual(await formatEdits(document), []);
	});

	test('formats a loose .bas file and keeps its Attribute line', async () => {
		const document = await openHidden(vscode.Uri.file(path.join(workspaceRoot(), 'Loose.bas')));
		assert.equal(document.languageId, 'vba');
		const edits = await formatEdits(document);
		const formatted = applyEdits(document.getText(), edits);
		assert.equal(formatted, [
			'Attribute VB_Name = "Loose"',
			'Option Explicit',
			'Sub hello()',
			'    Dim x As Long',
			'    If x = 1 Then',
			'        MsgBox "hi"',
			'    End If',
			'End Sub',
			'',
		].join('\r\n'));
	});

	test('Format All Modules formats and saves every module that needs it', async () => {
		await writeModule('BatchA', 'sub a()\r\nx=1\r\nend sub\r\n');
		await writeModule('BatchB', 'Sub b()\r\n    y = 2\r\nEnd Sub\r\n');

		await vscode.commands.executeCommand('xlide.formatAllModules', { filePath: workbookPath() }, { withoutPrompt: true });

		const onDisk = await until(() => {
			const modules = readModules(workbookPath());
			const a = modules.find((module) => module.name === 'BatchA')?.source;
			return a === 'Sub a()\r\n    x = 1\r\nEnd Sub\r\n' ? modules : undefined;
		}, 'BatchA should be formatted and saved into the workbook');
		assert.equal(onDisk.find((module) => module.name === 'BatchB')?.source, 'Sub b()\r\n    y = 2\r\nEnd Sub\r\n');
	});

	test('formats only the selected lines on Format Selection', async () => {
		const uri = await writeModule('FormatRange', 'sub a()\r\nx=1\r\nend sub\r\nsub b()\r\ny=2\r\nend sub\r\n');
		const document = await openHidden(uri);
		const edits = await vscode.commands.executeCommand<vscode.TextEdit[]>(
			'vscode.executeFormatRangeProvider', document.uri, new vscode.Range(3, 0, 5, 7), OPTIONS,
		) ?? [];
		assert.equal(applyEdits(document.getText(), edits), 'sub a()\r\nx=1\r\nend sub\r\nSub b()\r\n    y = 2\r\nEnd Sub\r\n');
	});
});
