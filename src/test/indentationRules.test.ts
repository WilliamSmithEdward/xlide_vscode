// The VBA indentation and Enter rules in a real VS Code: Reindent Lines lays
// out a module the way Format Document does, and Enter indents the next line
// as Format Document would. Smart Enter does not act on these lines (#Else,
// #ElseIf, and lines with a statement colon), so VS Code's rules decide alone.

import * as assert from 'node:assert/strict';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { activate, closeAllEditors, open, until, writeModule, workspaceRoot } from './support';

const indents = (text: string): string[] => text.split(/\r?\n/).map((line) => /^[ \t]*/.exec(line)![0]);

suite('VBA indentation rules', () => {
	suiteSetup(activate);
	teardown(async () => {
		await vscode.commands.executeCommand('workbench.action.revertAndCloseActiveEditor');
		await closeAllEditors();
	});

	test('Reindent Lines places headers with modifiers and directive arms where Format Document does', async () => {
		// Oracle, in all four Office hosts: modifier_type_enum_static_property_headers_*_compile.
		const layout = [
			'Public Type Rec',
			'    n As Long',
			'End Type',
			'Private Enum Colour',
			'    Red',
			'End Enum',
			'Public Static Property Get Total() As Long',
			'    Total = ThisWorkbook.Sheets.Count',
			'End Property',
			'Sub Main()',
			'    Dim x As Long',
			'    #If Win64 Then',
			'        x = ActiveDocument.Paragraphs.Count',
			'    #ElseIf VBA7 Then',
			'        x = ActivePresentation.Slides.Count',
			'    #Else',
			'        x = CurrentDb.TableDefs.Count',
			'    #End If',
			'End Sub',
			'',
		];
		const document = await open(await writeModule('IndentationRulesProbe', layout.map((line) => line.trimStart()).join('\r\n')));
		const editor = vscode.window.activeTextEditor!;
		editor.options = { insertSpaces: true, tabSize: 4 };
		const formatEdits = await vscode.commands.executeCommand<vscode.TextEdit[]>(
			'vscode.executeFormatDocumentProvider', document.uri, { tabSize: 4, insertSpaces: true });
		// Without an answer, the Format Document step below would pass on any layout.
		assert.ok(formatEdits && formatEdits.length > 0, 'Format Document should answer');
		await vscode.commands.executeCommand('editor.action.reindentlines');
		assert.deepEqual(indents(document.getText()), indents(layout.join('\r\n')));
		const reindented = document.getText();
		await vscode.commands.executeCommand('editor.action.formatDocument');
		assert.deepEqual(indents(document.getText()), indents(reindented), 'Format Document keeps the reindented layout');
	});

	test('Enter indents the next line after #Else and #ElseIf ... Then', async () => {
		const cases: Array<[line: string, indentAfter: string]> = [
			['    #Else', '        '],
			['    #ElseIf Win64 Then', '        '],
			["    #Else ' 32-bit", '        '],
			['    #ElseIf Win64', '    '],
		];
		const document = await open(await writeModule('IndentationEnterProbe', 'Sub Main()\r\nEnd Sub\r\n'));
		const editor = vscode.window.activeTextEditor!;
		editor.options = { insertSpaces: true, tabSize: 4 };
		// VS Code applies its indentation rules on Enter only to a line it can
		// tokenize at once: one at or before the first line still to be
		// tokenized. So only the lines from the second on are rewritten, and
		// Enter after `Case 1`, which indents under the old rules too, waits
		// until the first line is tokenized.
		const write = async (line: string): Promise<string> => {
			const rest = new vscode.Range(1, 0, document.lineCount, 0);
			assert.ok(await editor.edit((edit) => edit.replace(rest, `${line}\r\nEnd Sub\r\n`)));
			const end = document.lineAt(1).range.end;
			editor.selection = new vscode.Selection(end, end);
			await vscode.commands.executeCommand('type', { text: '\n' });
			return document.lineAt(2).text;
		};
		await until(async () => await write('    Case 1') === '        ' ? true : undefined,
			'Enter after `Case 1` should indent once the module is tokenized');
		const after: Array<[string, string]> = [];
		for (const [line] of cases) {
			after.push([line, await write(line)]);
		}
		assert.deepEqual(after, cases);
	});

	test('Reindent Lines places Else and ElseIf arms with a statement on their line where Format Document does', async () => {
		// Oracle, in all four Office hosts: block_if_else_same_line_statement_*_runtime
		// and block_if_elseif_same_line_statement_*_runtime.
		const layout = [
			'Sub Main()',
			'    Dim n As Long',
			'    If ThisWorkbook.Sheets.Count > 1 Then',
			'        n = 1',
			'    ElseIf ActiveDocument.Paragraphs.Count > 1 Then n = 2',
			'        n = n + 10',
			'    Else n = ActivePresentation.Slides.Count',
			'        n = n + CurrentDb.TableDefs.Count',
			'    End If',
			'End Sub',
			'',
		];
		const document = await open(await writeModule('IndentationArmsProbe', layout.map((line) => line.trimStart()).join('\r\n')));
		const editor = vscode.window.activeTextEditor!;
		editor.options = { insertSpaces: true, tabSize: 4 };
		const formatEdits = await vscode.commands.executeCommand<vscode.TextEdit[]>(
			'vscode.executeFormatDocumentProvider', document.uri, { tabSize: 4, insertSpaces: true });
		assert.ok(formatEdits && formatEdits.length > 0, 'Format Document should answer');
		await vscode.commands.executeCommand('editor.action.reindentlines');
		assert.deepEqual(indents(document.getText()), indents(layout.join('\r\n')));
		const reindented = document.getText();
		await vscode.commands.executeCommand('editor.action.formatDocument');
		assert.deepEqual(indents(document.getText()), indents(reindented), 'Format Document keeps the reindented layout');
	});

	test('Enter indents the next line after an Else or ElseIf arm with a statement, not after an ElseIf without Then', async () => {
		const cases: Array<[line: string, indentAfter: string]> = [
			['    Else n = 2', '        '],
			['    ElseIf n > 1 Then n = 2', '        '],
			['    Else: n = 2', '        '],
			['    ElseIf n > 1', '    '],
		];
		const document = await open(await writeModule('IndentationArmsEnterProbe', 'Sub Main()\r\nEnd Sub\r\n'));
		const editor = vscode.window.activeTextEditor!;
		editor.options = { insertSpaces: true, tabSize: 4 };
		// As in the #Else test above: Enter follows the rules only on a line
		// VS Code has tokenized, so wait for the module's first line.
		const write = async (line: string): Promise<string> => {
			const rest = new vscode.Range(1, 0, document.lineCount, 0);
			assert.ok(await editor.edit((edit) => edit.replace(rest, `${line}\r\nEnd Sub\r\n`)));
			const end = document.lineAt(1).range.end;
			editor.selection = new vscode.Selection(end, end);
			await vscode.commands.executeCommand('type', { text: '\n' });
			return document.lineAt(2).text;
		};
		await until(async () => await write('    Case 1') === '        ' ? true : undefined,
			'Enter after `Case 1` should indent once the module is tokenized');
		const after: Array<[string, string]> = [];
		for (const [line] of cases) {
			after.push([line, await write(line)]);
		}
		assert.deepEqual(after, cases);
	});

	test('Reindent Lines leaves the line after a one-line loop or a closed arm where Format Document does', async () => {
		// Oracle, Excel: issue447_07_runtime to issue447_12_runtime (one-line `For Each … In …: Next`)
		// and issue342_30_runtime (`For i = 0 To UBound(a): Next`) compile and run.
		const layout = [
			'Sub Main()',
			'    Dim n As Long',
			'    For Each ws In ThisWorkbook.Worksheets: Next',
			'    n = 1',
			'    For i = 1 To ActivePresentation.Slides.Count: n = n + i: Next',
			'    If ActiveDocument.Paragraphs.Count > 1 Then',
			'        n = 2',
			'    Else: End If',
			"    If n > CurrentDb.TableDefs.Count Then n = 0 ' only when it is greater Then",
			'    n = 3',
			'End Sub',
			'',
		];
		const document = await open(await writeModule('IndentationOneLineProbe', layout.map((line) => line.trimStart()).join('\r\n')));
		const editor = vscode.window.activeTextEditor!;
		editor.options = { insertSpaces: true, tabSize: 4 };
		const formatEdits = await vscode.commands.executeCommand<vscode.TextEdit[]>(
			'vscode.executeFormatDocumentProvider', document.uri, { tabSize: 4, insertSpaces: true });
		assert.ok(formatEdits && formatEdits.length > 0, 'Format Document should answer');
		await vscode.commands.executeCommand('editor.action.reindentlines');
		assert.deepEqual(indents(document.getText()), indents(layout.join('\r\n')));
		const reindented = document.getText();
		await vscode.commands.executeCommand('editor.action.formatDocument');
		assert.deepEqual(indents(document.getText()), indents(reindented), 'Format Document keeps the reindented layout');
	});

	test('Enter does not indent after a one-line loop, a closed arm or a Then inside a comment', async () => {
		// Lines Smart Enter leaves alone, so VS Code's rules decide; the open
		// openers it completes itself are covered by the unit tests.
		const cases: Array<[line: string, indentAfter: string]> = [
			['    For Each ws In ThisWorkbook.Worksheets: Next', '    '],
			['    While n < ActivePresentation.Slides.Count: n = n + 1: Wend', '    '],
			['    Else: End If', '    '],
			["    If n > CurrentDb.TableDefs.Count Then n = 0 ' only when it is greater Then", '    '],
			['    Else n = 2', '        '],
		];
		const document = await open(await writeModule('IndentationOneLineEnterProbe', 'Sub Main()\r\nEnd Sub\r\n'));
		const editor = vscode.window.activeTextEditor!;
		editor.options = { insertSpaces: true, tabSize: 4 };
		// As in the #Else test above: Enter follows the rules only on a line
		// VS Code has tokenized, so wait for the module's first line.
		const write = async (line: string): Promise<string> => {
			const rest = new vscode.Range(1, 0, document.lineCount, 0);
			assert.ok(await editor.edit((edit) => edit.replace(rest, `${line}\r\nEnd Sub\r\n`)));
			const end = document.lineAt(1).range.end;
			editor.selection = new vscode.Selection(end, end);
			await vscode.commands.executeCommand('type', { text: '\n' });
			return document.lineAt(2).text;
		};
		await until(async () => await write('    Case 1') === '        ' ? true : undefined,
			'Enter after `Case 1` should indent once the module is tokenized');
		const after: Array<[string, string]> = [];
		for (const [line] of cases) {
			after.push([line, await write(line)]);
		}
		assert.deepEqual(after, cases);
	});

	test('Reindent Lines leaves the line after a one-line procedure, Type or Enum where Format Document does, and indents after Do:', async () => {
		// Oracle, in all four Office hosts: one_line_procedures_types_and_do_colon_*_runtime.
		const layout = [
			'Private Type Rec: n As Long: End Type',
			'Public Enum Colour: Red = 1: Green: End Enum',
			'Public Sub Delete(): End Sub',
			'Private Function Twice(ByVal v As Long) As Long: Twice = 2 * v: End Function',
			'Public Static Property Get Count() As Long: Count = ThisWorkbook.Sheets.Count: End Property',
			'Sub Main()',
			'    Dim n As Long',
			'    Do: n = n + 1',
			'        If n > ActiveDocument.Paragraphs.Count Then Exit Do',
			'    Loop',
			'    n = ActivePresentation.Slides.Count + CurrentDb.TableDefs.Count',
			'End Sub',
			'',
		];
		const document = await open(await writeModule('IndentationOneLineProcProbe', layout.map((line) => line.trimStart()).join('\r\n')));
		const editor = vscode.window.activeTextEditor!;
		editor.options = { insertSpaces: true, tabSize: 4 };
		const formatEdits = await vscode.commands.executeCommand<vscode.TextEdit[]>(
			'vscode.executeFormatDocumentProvider', document.uri, { tabSize: 4, insertSpaces: true });
		assert.ok(formatEdits && formatEdits.length > 0, 'Format Document should answer');
		await vscode.commands.executeCommand('editor.action.reindentlines');
		assert.deepEqual(indents(document.getText()), indents(layout.join('\r\n')));
		const reindented = document.getText();
		await vscode.commands.executeCommand('editor.action.formatDocument');
		assert.deepEqual(indents(document.getText()), indents(reindented), 'Format Document keeps the reindented layout');
	});

	test('Enter does not indent after a one-line procedure, Type or Enum, and indents after Do:', async () => {
		// Lines Smart Enter leaves alone, so VS Code's rules decide.
		const cases: Array<[line: string, indentAfter: string]> = [
			['Public Sub Delete(): End Sub', ''],
			['Private Function Twice(ByVal v As Long) As Long: Twice = 2 * v: End Function', ''],
			['Property Let Count(ByVal v As Long): m = v: End Property', ''],
			['Private Type Rec: n As Long: End Type', ''],
			['Public Enum Colour: Red = 1: Green: End Enum', ''],
			['Sub Main(): n = ThisWorkbook.Sheets.Count', '    '],
			['    Do: n = n + 1', '        '],
		];
		const document = await open(await writeModule('IndentationOneLineProcEnter', 'Sub Main()\r\nEnd Sub\r\n'));
		const editor = vscode.window.activeTextEditor!;
		editor.options = { insertSpaces: true, tabSize: 4 };
		// As in the #Else test above: Enter follows the rules only on a line
		// VS Code has tokenized, so wait for the module's first line.
		const write = async (line: string): Promise<string> => {
			const rest = new vscode.Range(1, 0, document.lineCount, 0);
			assert.ok(await editor.edit((edit) => edit.replace(rest, `${line}\r\nEnd Sub\r\n`)));
			const end = document.lineAt(1).range.end;
			editor.selection = new vscode.Selection(end, end);
			await vscode.commands.executeCommand('type', { text: '\n' });
			return document.lineAt(2).text;
		};
		await until(async () => await write('    Case 1') === '        ' ? true : undefined,
			'Enter after `Case 1` should indent once the module is tokenized');
		const after: Array<[string, string]> = [];
		for (const [line] of cases) {
			after.push([line, await write(line)]);
		}
		assert.deepEqual(after, cases);
	});

	test('Reindent Lines indents non-ASCII headers and loop variables in every host example', async () => {
		const layout = [
			'Public Type État', '    n As Long', 'End Type',
			'Private Enum Größe', '    Small', 'End Enum',
			'Public Static Property Get Länge() As Long', '    Länge = 1', 'End Property',
			'Sub 名前()', '    Dim ячейка As Object',
			...['ThisWorkbook.Worksheets', 'ActiveDocument.Paragraphs',
				'ActivePresentation.Slides', 'CurrentDb.TableDefs'].flatMap((collection) => [
				`    For Each ячейка In ${collection}`, '        n = n + 1', '    Next',
			]),
			'    For État = 1 To 5', '        n = n + État', '    Next', 'End Sub', '',
		];
		const uri = vscode.Uri.file(path.join(workspaceRoot(), 'NamesIndentProbe.bas'));
		await vscode.workspace.fs.writeFile(uri, Buffer.from(layout.map((line) => line.trimStart()).join('\r\n')));
		const document = await open(uri);
		vscode.window.activeTextEditor!.options = { insertSpaces: true, tabSize: 4 };
		await vscode.commands.executeCommand('editor.action.reindentlines');
		assert.deepEqual(indents(document.getText()), indents(layout.join('\r\n')));
	});

	test('Reindent Lines reads a keyword joined to a non-ASCII character as a name', async () => {
		const layout = [
			'Sub Main()',
			'    If ÉRem > ThisWorkbook.Sheets.Count Then', '        n = 1', '    End If',
			'    If Thené = ActiveDocument.Paragraphs.Count Then', '        n = 2', '    End If',
			'    Nexté = 3', '    Caseé = 4', '    Elseé = 5', '    Loopé = 6', '    Wendé = 7',
			'    With ÉRem: End With',
			'    While n < ActivePresentation.Slides.Count: Wendé = 8', '        n = n + 1', '    Wend',
			'    Do While CurrentDb.TableDefs.Count > n: Loopé = 9', '        n = n + 1', '    Loop',
			'End Sub', '',
		];
		const uri = vscode.Uri.file(path.join(workspaceRoot(), 'KeywordNamesIndentProbe.bas'));
		await vscode.workspace.fs.writeFile(uri, Buffer.from(layout.map((line) => line.trimStart()).join('\r\n')));
		const document = await open(uri);
		vscode.window.activeTextEditor!.options = { insertSpaces: true, tabSize: 4 };
		await vscode.commands.executeCommand('editor.action.reindentlines');
		assert.deepEqual(indents(document.getText()), indents(layout.join('\r\n')));
	});

	test('Enter indents after a non-ASCII procedure header', async () => {
		const uri = vscode.Uri.file(path.join(workspaceRoot(), 'NamesEnterProbe.bas'));
		await vscode.workspace.fs.writeFile(uri, Buffer.from('Sub Main()\r\nEnd Sub\r\n'));
		const document = await open(uri);
		const editor = vscode.window.activeTextEditor!;
		editor.options = { insertSpaces: true, tabSize: 4 };
		const write = async (line: string): Promise<string> => {
			await until(async () => await editor.edit((edit) => edit.replace(new vscode.Range(1, 0, document.lineCount, 0),
				`${line}\r\nEnd Sub\r\n`)) || undefined, 'Replace the probe after pending editor edits settle');
			const end = document.lineAt(1).range.end;
			editor.selection = new vscode.Selection(end, end);
			await vscode.commands.executeCommand('type', { text: '\n' });
			return document.lineAt(2).text;
		};
		await until(async () => await write('    Case 1') === '        ' ? true : undefined,
			'Enter waits for the module to be tokenized');
		for (const line of ['Sub 名前()', 'Function État() As Long', 'Public Static Property Get Größe()']) {
			assert.equal(await write(line), '    ', line);
		}
	});

});
