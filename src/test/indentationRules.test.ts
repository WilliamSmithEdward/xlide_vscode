// The VBA indentation and Enter rules in a real VS Code: Reindent Lines lays
// out a module the way Format Document does, and Enter after a line indents
// the next one as the block reading of the formatter says. Smart Enter does
// not act on these lines (none of them is a complete block opener on its own
// line), so VS Code's rules decide alone.

import * as assert from 'node:assert/strict';
import * as vscode from 'vscode';
import { activate, closeAllEditors, open, until, writeModule } from './support';

const indents = (text: string): string[] => text.split(/\r?\n/).map((line) => /^[ \t]*/.exec(line)![0]);

suite('VBA indentation rules', () => {
	suiteSetup(activate);
	teardown(async () => {
		await vscode.commands.executeCommand('workbench.action.revertAndCloseActiveEditor');
		await closeAllEditors();
	});

	test('Reindent Lines places these block forms where Format Document does', async () => {
		// Oracle, in all four Office hosts: modifier_type_enum_static_property_headers_*_compile,
		// block_if_else(if)_same_line_statement_*_runtime and
		// one_line_loop_and_with_blocks_*_compile (the For, Do, While and With lines).
		const layout = [
			'Public Type Rec',
			'    n As Long',
			'End Type',
			'Private Enum Colour',
			'    Red',
			'End Enum',
			'Public Static Property Get Total() As Long',
			'    Total = 1',
			'End Property',
			'Public Sub Delete(): End Sub',
			'Sub Main()',
			'    Dim i As Long, x As Long',
			'    #If Win64 Then',
			'        x = 64',
			'    #ElseIf VBA7 Then',
			'        x = 7',
			'    #Else',
			'        x = 0',
			'    #End If',
			'    If x = 0 Then',
			'        x = 1',
			"    ElseIf x = 1 Then x = 2 ' only if x = 1 Then",
			'        x = 3',
			'    Else x = ThisWorkbook.Sheets.Count',
			'        x = 4',
			'    End If',
			'    For i = 1 To 5: x = x + 1: Next',
			'    For Each p In ActiveDocument.Paragraphs: x = x + 1: Next',
			'    Do While x > 0: x = x - 1: Loop',
			'    With ActivePresentation.Slides(1): .Name = "First": End With',
			'    While x < 0: Wend',
			'    x = CurrentDb.TableDefs.Count',
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

	test('Enter indents the next line by the block reading of the line', async () => {
		const cases: Array<[line: string, indentAfter: string]> = [
			['    #Else', '        '],
			['    #ElseIf Win64 Then', '        '],
			['    Else x = ThisWorkbook.Sheets.Count', '        '],
			['    ElseIf ActiveDocument.Paragraphs.Count > 1 Then n = 2', '        '],
			['    Case Else', '        '],
			["    If ready Then n = 1 ' only if ready Then", '    '],
			['    For Each t In CurrentDb.TableDefs: n = n + 1: Next', '    '],
			['    Do While n > 0: n = n - 1: Loop', '    '],
			['    With ActivePresentation.Slides(1): .Name = "x": End With', '    '],
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
});
