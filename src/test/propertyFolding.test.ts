// Folding in a real VS Code, through XLIDE's folding provider, for a module
// written for each Office host, indented and flat. A Property with Static among
// its modifiers folds to its End Property.

import * as assert from 'node:assert/strict';
import * as vscode from 'vscode';
import { activate, closeAllEditors, open, writeModule } from './support';

const propertyModule = (collection: string, indent: string): string => [
	'Option Explicit',
	'Private mCount As Long',
	'Public Static Property Get Count() As Long',
	`${indent}Count = ${collection}.Count + mCount`,
	'End Property',
	'Static Public Property Let Count(ByVal v As Long)',
	`${indent}mCount = v`,
	'End Property',
	'',
].join('\r\n');

const foldRanges = async (uri: vscode.Uri): Promise<Array<[number, number]>> => {
	const ranges = await vscode.commands.executeCommand<vscode.FoldingRange[]>('vscode.executeFoldingRangeProvider', uri);
	return (ranges ?? []).map((range): [number, number] => [range.start + 1, range.end + 1]).sort((a, b) => a[0] - b[0] || a[1] - b[1]);
};

suite('VBA folding', () => {
	suiteSetup(activate);
	teardown(async () => {
		await vscode.commands.executeCommand('workbench.action.revertAndCloseActiveEditor');
		await closeAllEditors();
	});

	test('each Property folds from its header to End Property, in a module written for each Office host', async () => {
		// Oracle: issue1496_{,word_,powerpoint_,access_}static_public_property_compile.
		const hosts = {
			Excel: 'ThisWorkbook.Worksheets',
			Word: 'ActiveDocument.Paragraphs',
			PowerPoint: 'ActivePresentation.Slides',
			Access: 'CurrentDb.TableDefs',
		};
		for (const [host, collection] of Object.entries(hosts)) {
			for (const [layout, indent] of [['indented', '    '], ['flat', '']]) {
				const document = await open(await writeModule(`StaticPropFold${host}${layout === 'flat' ? 'Flat' : ''}`, propertyModule(collection, indent)));
				assert.deepEqual(await foldRanges(document.uri), [[3, 5], [6, 8]], `${host}, ${layout}`);
			}
		}
	});

	test('If, loop, With, Select Case and #If blocks fold whole, and their arms up to the next, for each Office host (issue #1576)', async () => {
		const hosts = {
			Excel: 'ThisWorkbook.Worksheets',
			Word: 'ActiveDocument.Paragraphs',
			PowerPoint: 'ActivePresentation.Slides',
			Access: 'CurrentDb.TableDefs',
		};
		const blocks = (collection: string): string[] => [
			'Sub Main()',
			`If ${collection}.Count > 1 Then`,
			'n = 1',
			'Else',
			'n = 2',
			'End If',
			`For Each item In ${collection}`,
			'n = n + 1',
			'Next',
			'Select Case n',
			'Case 1',
			'n = 2',
			'Case Else',
			'n = 3',
			'End Select',
			`With ${collection}`,
			'n = .Count',
			'End With',
			'#If VBA7 Then',
			'n = 4',
			'#Else',
			'n = 5',
			'#End If',
			'End Sub',
			'',
		];
		for (const [host, collection] of Object.entries(hosts)) {
			const document = await open(await writeModule(`BlockFold${host}`, blocks(collection).join('\r\n')));
			assert.deepEqual(await foldRanges(document.uri), [
				[1, 24], [2, 6], [4, 5], [7, 9], [10, 15], [11, 12], [13, 14], [16, 18], [19, 23], [21, 22],
			], host);
		}
	});
});
