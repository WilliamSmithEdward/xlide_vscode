// Cross-application references through a real VS Code: code that names
// another application early-bound is reported when the project does not
// reference that application's library, and the quick fix writes the
// reference into the project rather than sending the user to the VBE.
//
// The suite works on its own copy of the fixture workbook. Adding a reference
// changes which object models every module in a project resolves against, so
// doing it to the shared copy would move the ground under the other suites.

import * as assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { encodeModuleUri } from '../xlideFileSystem';
import { listReferences, writeModule } from '../vba/projectService';
import { activate, closeAllEditors, open, until, workbookPath, workspaceRoot } from './support';

const MODULE = 'InteropProbe';
const SOURCE = [
	'Option Explicit',
	'',
	'Public Sub CopyToWord()',
	'    Dim doc As Word.Document',
	'    Set doc = Nothing',
	'End Sub',
	'',
].join('\r\n');

suite('Project references', () => {
	let probe: string;

	suiteSetup(async () => {
		await activate();
		probe = path.join(workspaceRoot(), 'InteropProbe.xlsm');
		fs.copyFileSync(workbookPath(), probe);
		writeModule(probe, MODULE, SOURCE, 'standard');
	});

	teardown(async () => {
		await closeAllEditors();
	});

	test('reports a library the project does not reference, and the fix adds it', async () => {
		assert.ok(
			!listReferences(probe).some((one) => one.name === 'Word'),
			'the workbook should start without a Word reference',
		);
		const document = await open(encodeModuleUri(probe, MODULE));

		const diagnostic = await until(
			() => vscode.languages.getDiagnostics(document.uri)
				.find((one) => one.code === 'missing-library-reference'),
			'the missing reference should be reported',
		);
		assert.match(diagnostic.message, /'Word' is not referenced by this project/);
		assert.equal(document.getText(diagnostic.range), 'Word.Document');

		const actions = await vscode.commands.executeCommand<vscode.CodeAction[]>(
			'vscode.executeCodeActionProvider', document.uri, diagnostic.range,
		) ?? [];
		const fix = actions.find((one) => one.title === 'Add a reference to the Word object library');
		assert.ok(fix?.command, `quick fixes offered: ${actions.map((one) => one.title).join(', ')}`);

		await vscode.commands.executeCommand(fix.command.command, ...fix.command.arguments ?? []);

		const written = listReferences(probe).find((one) => one.name === 'Word');
		assert.ok(written, 'the fix should write the reference into the project');
		assert.match(written.libid, /\{00020905-0000-0000-C000-000000000046\}/);

		// The reference is what the analyzer resolves `Word.Document` against,
		// so the finding should go on its own once the project carries it.
		await until(
			() => vscode.languages.getDiagnostics(document.uri)
				.every((one) => one.code !== 'missing-library-reference') || undefined,
			'the finding should clear once the project references Word',
		);
	});
});
