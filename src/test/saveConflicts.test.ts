// Saving over a change, in a real VS Code. VS Code stops a save at "File
// Modified Since" only when the file's size differs as well as its mtime, and
// XLIDE moved only the mtime after its own writes and for a module it could
// not read. So a save went straight over an agent's write, and a module
// deleted in the VBE came back when its editor was saved.

import * as assert from 'node:assert/strict';
import * as vscode from 'vscode';
import { deleteModule as deleteOutsideXlide, listModules, readModules } from '../vba/projectService';
import { encodeFormMarkupUri } from '../xlideFileSystem';
import { activate, closeAllEditors, moduleUri, open, workbookPath, writeModule as writeOutsideXlide } from './support';

/** Writes a module the way an agent does: through XLIDE's write tool. */
async function agentWrite(moduleName: string, source: string): Promise<void> {
	await vscode.lm.invokeTool('xlide_writeModule', {
		input: { filePath: workbookPath(), moduleName, source },
		toolInvocationToken: undefined,
	}, new vscode.CancellationTokenSource().token);
}

/** Longer than the watcher's settle time and the comparison after it. */
function afterTheChangeIsNoticed(): Promise<void> {
	return new Promise((resolve) => setTimeout(resolve, 1500));
}

function moduleSource(moduleName: string): string | undefined {
	return readModules(workbookPath()).find((module) => module.name === moduleName)?.source;
}

/** Types a line at the end of the open module, leaving it with unsaved edits. */
async function typeLine(text: string): Promise<void> {
	const editor = vscode.window.activeTextEditor!;
	await editor.edit((edit) => edit.insert(new vscode.Position(editor.document.lineCount - 1, 0), `${text}\r\n`));
}

suite('Saving over a change', () => {
	suiteSetup(async () => {
		await activate();
	});

	teardown(async () => {
		await vscode.commands.executeCommand('workbench.action.files.revert');
		await closeAllEditors();
	});

	for (const savedOnce of [false, true]) {
		test(`stops at the conflict when an agent wrote the module${savedOnce ? ', after a save' : ''}`, async () => {
			const name = savedOnce ? 'ConflictAgentSaved' : 'ConflictAgent';
			await agentWrite(name, 'Option Explicit\r\n\r\nSub A()\r\nEnd Sub\r\n');
			const document = await open(moduleUri(name));
			if (savedOnce) {
				await typeLine("' saved once");
				assert.equal(await document.save(), true);
			}
			await typeLine("' unsaved");

			await agentWrite(name, 'Option Explicit\r\n\r\nSub A()\r\n    agentWasHere = 1\r\nEnd Sub\r\n');

			assert.equal(await document.save(), false, 'the save should stop at the conflict');
			assert.ok(moduleSource(name)?.includes('agentWasHere'), "the agent's change should still be in the workbook");
		});
	}

	test('stops at the conflict when the module was deleted outside XLIDE', async () => {
		await agentWrite('ConflictGone', 'Option Explicit\r\n\r\nSub G()\r\nEnd Sub\r\n');
		const document = await open(moduleUri('ConflictGone'));
		await typeLine("' saved once");
		assert.equal(await document.save(), true);
		await typeLine("' unsaved");

		deleteOutsideXlide(workbookPath(), 'ConflictGone');
		await afterTheChangeIsNoticed();

		assert.equal(await document.save(), false, 'the save should stop at the conflict');
		assert.equal(moduleSource('ConflictGone'), undefined, 'the deleted module should stay deleted');
		// Reverting an editor whose module is gone fails inside VS Code, so the
		// module comes back before the teardown reverts it.
		writeOutsideXlide('ConflictGone', 'Option Explicit\r\n');
	});

	test('saves without a conflict after its own save the engine stored differently, when the change was elsewhere', async () => {
		// The engine drops blank lines above a module's code. A save was
		// recorded as the text written, so the next outside change anywhere in
		// the workbook stopped this module's next save at a conflict.
		await agentWrite('StoredDifferently', 'Option Explicit\r\n');
		const document = await open(moduleUri('StoredDifferently'));
		await vscode.window.activeTextEditor!.edit((edit) => edit.replace(
			new vscode.Range(0, 0, document.lineCount, 0),
			'\r\n\r\nOption Explicit\r\nSub A()\r\nEnd Sub\r\n',
		));
		assert.equal(await document.save(), true);
		await typeLine("' unsaved");

		writeOutsideXlide('StoredDifferentlyOther', 'Option Explicit\r\n');
		await afterTheChangeIsNoticed();

		assert.equal(await document.save(), true, 'a change to another module is no conflict');
	});

	test('saves a form designer without a conflict after its own reformatted save, when the change was elsewhere', async () => {
		const form = listModules(workbookPath()).find((module) => module.type === 'userform');
		assert.ok(form, 'the fixture should have a UserForm');
		const document = await open(encodeFormMarkupUri(workbookPath(), form.name));
		const original = document.getText();
		// Only the layout changes: the designer writes its markup back its own way.
		await vscode.window.activeTextEditor!.edit((edit) => edit.replace(
			new vscode.Range(0, 0, document.lineCount, 0),
			original.replace(/<(\w+) /, '<$1  '),
		));
		assert.equal(await document.save(), true);
		await vscode.window.activeTextEditor!.edit((edit) => edit.insert(new vscode.Position(0, 0), ' '));

		writeOutsideXlide('FormElsewhere', 'Option Explicit\r\n');
		await afterTheChangeIsNoticed();

		assert.equal(await document.save(), true, 'a change to another module is no conflict');
	});

	test('stops at the conflict over an outside change that kept the byte length', async () => {
		await agentWrite('ConflictSameLength', 'Option Explicit\r\n\r\nSub Mine()\r\nEnd Sub\r\n');
		const document = await open(moduleUri('ConflictSameLength'));
		await typeLine("' unsaved");

		writeOutsideXlide('ConflictSameLength', 'Option Explicit\r\n\r\nSub Them()\r\nEnd Sub\r\n');
		await afterTheChangeIsNoticed();

		assert.equal(await document.save(), false, 'the save should stop at the conflict');
		assert.ok(moduleSource('ConflictSameLength')?.includes('Sub Them()'));
	});
});
