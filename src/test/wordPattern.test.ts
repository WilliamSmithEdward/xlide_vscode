// The language configuration's word pattern in a real VS Code. Quick suggestions
// open only on a word at the cursor, and word-based suggestions, occurrence
// highlighting and the Ctrl+hover link read the same pattern, so the letters of
// a number must not be a word. Extensions cannot read the pattern: without a
// pattern of its own, TextDocument.getWordRangeAtPosition uses VS Code's default.
// Expand Selection runs in the editor, on the model's own words.

import * as assert from 'node:assert/strict';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { activate, closeAllEditors, open, workspaceRoot, writeModule } from './support';

const LONG_NAME = `VeryLong${'Name'.repeat(20)}`;
const LINES = [
	'Sub Main()',
	'    n = 1.5e3 + &HFF + 2D-1 + 1.e3',
	'    v = rng.e3 + hex1',
	`    ${LONG_NAME} = 1`,
	'End Sub',
	'',
];

suite('VBA word pattern', () => {
	suiteSetup(activate);
	teardown(closeAllEditors);

	test('finds no word inside a number, and the whole of a name', async () => {
		const document = await open(await writeModule('WordPatternProbe', LINES.join('\r\n')));
		/** The text of each range Expand Selection steps through from `offset` into `text`. */
		const steps = async (line: number, text: string, offset: number): Promise<string[]> => {
			const position = new vscode.Position(line, document.lineAt(line).text.indexOf(text) + offset);
			const [chain] = await vscode.commands.executeCommand<vscode.SelectionRange[]>(
				'vscode.executeSelectionRangeProvider', document.uri, [position],
			);
			const texts: string[] = [];
			for (let range: vscode.SelectionRange | undefined = chain; range; range = range.parent) {
				texts.push(document.getText(range.range));
			}
			return texts;
		};
		for (const [number, letters] of [['1.5e3', 'e3'], ['&HFF', 'HFF'], ['2D-1', 'D'], ['1.e3', 'e3']]) {
			const texts = await steps(1, number, number.indexOf(letters) + 1);
			assert.ok(!texts.includes(letters), `${number}: ${JSON.stringify(texts.slice(0, 3))}`);
		}
		assert.ok((await steps(2, 'e3', 1)).includes('e3'), 'a member spelled like an exponent is a word');
		assert.ok((await steps(2, 'hex1', 2)).includes('hex1'), 'a name ending in a digit is a word');
		// VS Code looks for the word in growing windows: a start inside a long name is refused, then found.
		assert.ok((await steps(3, LONG_NAME, LONG_NAME.length - 1)).includes(LONG_NAME), 'a long name is one word');
	});

	// A module written through the host keeps only its code page's characters, so this is a loose file.
	test('finds the whole of a name outside ASCII (issue #1604)', async () => {
		const names = ['値', 'Ünïcödé', 'ค่า', 'cafe\u0301'];
		const uri = vscode.Uri.file(path.join(workspaceRoot(), 'WordPatternNames.bas'));
		await vscode.workspace.fs.writeFile(uri, Buffer.from(`Sub Main()\r\n    ${names.join(' + ')}\r\nEnd Sub\r\n`));
		try {
			const document = await open(uri);
			for (const name of names) {
				const position = new vscode.Position(1, document.lineAt(1).text.indexOf(name) + 1);
				const [chain] = await vscode.commands.executeCommand<vscode.SelectionRange[]>(
					'vscode.executeSelectionRangeProvider', document.uri, [position],
				);
				const texts: string[] = [];
				for (let range: vscode.SelectionRange | undefined = chain; range; range = range.parent) {
					texts.push(document.getText(range.range));
				}
				assert.ok(texts.includes(name), `${name}: ${JSON.stringify(texts.slice(0, 3))}`);
			}
		} finally {
			await vscode.workspace.fs.delete(uri);
		}
	});
});
