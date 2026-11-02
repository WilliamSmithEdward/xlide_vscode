// A Property with Static among its modifiers folds in a real VS Code. XLIDE
// registers no folding provider, so VS Code folds a module by indentation and
// by the language configuration's start and end markers; a Property header the
// start marker misses folds only by indentation, without its End Property.

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

suite('VBA folding of a Static Property', () => {
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
});
