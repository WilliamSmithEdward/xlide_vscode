// The agent tools that work on parts of a module and on a folder of module
// files, in a real VS Code against the fixture workbook: the parts come back
// numbered as the module reads, the edits land where the result says, and an
// import applies what the folder holds and no more.

import * as assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { readModules } from '../vba/projectService';
import { activate, workbookPath } from './support';

async function invoke(tool: string, input: Record<string, unknown>): Promise<string> {
	const result = await vscode.lm.invokeTool(tool, {
		input: { filePath: workbookPath(), ...input },
		toolInvocationToken: undefined,
	}, new vscode.CancellationTokenSource().token);
	return result.content
		.map((part) => (part instanceof vscode.LanguageModelTextPart ? part.value : ''))
		.join('');
}

function source(moduleName: string): string | undefined {
	return readModules(workbookPath()).find((module) => module.name === moduleName)?.source;
}

suite('Editing parts of a module through the agent tools', () => {
	suiteSetup(async () => {
		await activate();
	});

	test('reads the procedures asked for, and edits them where the read numbered them', async () => {
		await invoke('xlide_writeModule', {
			moduleName: 'PartsProbe',
			source: [
				'Option Explicit',
				'',
				"' Adds one",
				'Public Function Inc(n As Long) As Long',
				'    Inc = n + 1',
				'End Function',
				'',
				'Public Sub Tail()',
				'End Sub',
				'',
			].join('\r\n'),
		});
		const read = await invoke('xlide_readModule', {
			moduleName: 'PartsProbe',
			procedures: ['Inc'],
			ranges: [{ startLine: 1, endLine: 1 }],
		});
		const token = /contentToken: (\S+) \((\d+) lines\)/.exec(read);
		assert.ok(token, read);
		assert.equal(token[2], '9');
		assert.ok(read.includes('--- lines 1-1\nOption Explicit'), read);
		assert.ok(read.includes("--- Function Inc (lines 3-6)\n' Adds one\nPublic Function Inc(n As Long) As Long\n    Inc = n + 1\nEnd Function"), read);

		const edited = await invoke('xlide_editModule', {
			moduleName: 'PartsProbe',
			expectedContentToken: token[1],
			edits: [
				{ procedure: 'Inc', text: "' Adds two\r\nPublic Function Inc(n As Long) As Long\r\n    Inc = n + 2\r\nEnd Function" },
				{ insertAfterLine: 1, text: 'Private m As Long' },
			],
		});
		assert.ok(edited.includes('- Function Inc: now lines 4-7'), edited);
		assert.ok(edited.includes('- after line 1: now lines 2-2'), edited);
		assert.equal(source('PartsProbe'), [
			'Option Explicit',
			'Private m As Long',
			'',
			"' Adds two",
			'Public Function Inc(n As Long) As Long',
			'    Inc = n + 2',
			'End Function',
			'',
			'Public Sub Tail()',
			'End Sub',
			'',
		].join('\r\n'));

		// The read's token is stale now; the result carries the new one.
		const stale = await invoke('xlide_editModule', {
			moduleName: 'PartsProbe',
			expectedContentToken: token[1],
			edits: [{ insertAfterLine: 0, text: "' x" }],
		});
		assert.ok(stale.includes('changed since it was read'), stale);
		const fresh = /contentToken: (\S+)/.exec(edited);
		assert.ok(fresh, edited);
		const again = await invoke('xlide_editModule', {
			moduleName: 'PartsProbe',
			expectedContentToken: fresh[1],
			edits: [{ startLine: 2, endLine: 2, text: '' }],
		});
		assert.ok(again.includes('- lines 2-2: removed'), again);
		assert.ok(!source('PartsProbe')?.includes('Private m As Long'));
	});
});

suite('Importing modules through the agent tool', () => {
	let folder: string;

	suiteSetup(async () => {
		await activate();
		folder = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-import-'));
	});

	suiteTeardown(() => {
		fs.rmSync(folder, { recursive: true, force: true });
	});

	test('applies the folder: one module updated, one created, one left as it is', async () => {
		await invoke('xlide_writeModule', { moduleName: 'ImportOld', source: 'Sub Old()\r\nEnd Sub\r\n' });
		await invoke('xlide_writeModule', { moduleName: 'ImportSame', source: 'Sub Same()\r\nEnd Sub\r\n' });
		// The folder as an export leaves it, then edited the way a developer
		// edits it: one file changed, one file new.
		await invoke('xlide_exportModules', { exportFolder: folder });
		const old = path.join(folder, 'ImportOld.bas');
		fs.writeFileSync(old, fs.readFileSync(old, 'utf8').replace('End Sub', '    Debug.Print 1\r\nEnd Sub'));
		fs.writeFileSync(path.join(folder, 'ImportNew.bas'), 'Sub Fresh()\r\nEnd Sub\r\n');

		const report = JSON.parse(await invoke('xlide_importModules', {
			importFolder: folder,
			modules: ['ImportOld', 'ImportSame', 'ImportNew'],
		})) as { importFolder: string; updated: string[]; created: string[]; skipped: unknown[]; failed: string[] };

		assert.equal(report.importFolder, folder);
		assert.deepEqual(report.updated, ['ImportOld']);
		assert.deepEqual(report.created, ['ImportNew']);
		assert.deepEqual(report.skipped, [{ module: 'ImportSame', file: 'ImportSame.bas', reason: 'unchanged' }]);
		assert.deepEqual(report.failed, []);
		assert.equal(source('ImportOld'), 'Sub Old()\r\n    Debug.Print 1\r\nEnd Sub\r\n');
		assert.equal(source('ImportNew'), 'Sub Fresh()\r\nEnd Sub\r\n');
	});

	test('takes the folder the export recorded, and refuses a name the folder has no file for', async () => {
		// ImportSame's file is the export's own, header and all, so it still
		// reads as unchanged; a file written without a header does not.
		const report = JSON.parse(await invoke('xlide_importModules', { modules: ['ImportSame'] })) as { importFolder: string; skipped: Array<{ reason: string }> };
		assert.equal(report.importFolder, folder);
		assert.deepEqual(report.skipped.map((skip) => skip.reason), ['unchanged']);

		const refused = await invoke('xlide_importModules', { importFolder: folder, modules: ['Nope'] });
		assert.ok(refused.startsWith(`No file in ${folder} is for "Nope".`), refused);
	});
});
