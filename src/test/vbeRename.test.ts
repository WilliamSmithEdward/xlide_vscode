// A module renamed outside XLIDE - in the VBE, and saved - while an editor is
// open on it, in a real VS Code. The suite renames with the engine directly,
// which is exactly that: the extension did not make the change.
//
// Measured before XLIDE followed such a rename: the tree refreshed, and the
// editor stayed on the old name, marked deleted, showing a module the project
// no longer had.

import * as assert from 'node:assert/strict';
import * as vscode from 'vscode';
import { deleteModule, renameModule } from '../vba/projectService';
import { activate, closeAllEditors, moduleUri, open, until, workbookPath } from './support';

/** Writes a module through XLIDE, so the extension knows it is there. */
async function agentWrite(moduleName: string, source: string): Promise<void> {
	await vscode.lm.invokeTool('xlide_writeModule', {
		input: { filePath: workbookPath(), moduleName, source },
		toolInvocationToken: undefined,
	}, new vscode.CancellationTokenSource().token);
}

function tabsOn(moduleName: string): vscode.Tab[] {
	const uri = moduleUri(moduleName).toString();
	return vscode.window.tabGroups.all
		.flatMap((group) => group.tabs)
		.filter((tab) => tab.input instanceof vscode.TabInputText && tab.input.uri.toString() === uri);
}

/** Every tab, for a failure message that shows what the editor did instead. */
function describeTabs(): string {
	return vscode.window.tabGroups.all
		.flatMap((group) => group.tabs)
		.map((tab) => `${tab.label}${tab.isDirty ? ' (dirty)' : ''}${tab.isActive ? ' (active)' : ''}`)
		.join(', ') || 'no tabs';
}

/** Longer than the watcher's settle time and the comparison after it. */
function afterTheChangeIsNoticed(): Promise<void> {
	return new Promise((resolve) => setTimeout(resolve, 2500));
}

suite('A module renamed outside XLIDE', () => {
	suiteSetup(async () => {
		await activate();
	});

	teardown(async () => {
		// Revert throws for an editor whose module is gone, which is what a
		// failing check leaves behind.
		await Promise.resolve(vscode.commands.executeCommand('workbench.action.files.revert')).catch(() => undefined);
		await closeAllEditors();
	});

	test('takes its editor to the new name, at the same place', async () => {
		const source = 'Option Explicit\r\n\r\nPublic Sub First()\r\nEnd Sub\r\n\r\nPublic Sub Second()\r\nEnd Sub\r\n';
		await agentWrite('VbeRenameOld', source);
		await open(moduleUri('VbeRenameOld'));
		const selection = new vscode.Selection(5, 4, 5, 10);
		vscode.window.activeTextEditor!.selection = selection;

		renameModule(workbookPath(), 'VbeRenameOld', 'VbeRenameNew');

		await until(
			() => (tabsOn('VbeRenameOld').length === 0 && tabsOn('VbeRenameNew').length === 1 ? true : undefined),
			'the editor should follow the module to its new name',
			8000,
		).catch((err: Error) => {
			throw new Error(`${err.message}; tabs now: ${describeTabs()}`);
		});
		const editor = vscode.window.visibleTextEditors.find((candidate) => candidate.document.uri.toString() === moduleUri('VbeRenameNew').toString());
		assert.ok(editor, 'the new name should be showing');
		assert.equal(editor.document.getText(), source);
		assert.deepEqual(editor.selection, selection);
	});

	test('leaves an editor with unsaved edits where it is', async () => {
		await agentWrite('VbeDirtyOld', 'Option Explicit\r\n\r\nSub Keep()\r\nEnd Sub\r\n');
		const document = await open(moduleUri('VbeDirtyOld'));
		await vscode.window.activeTextEditor!.edit((edit) => edit.insert(new vscode.Position(4, 0), "' mine\r\n"));
		assert.equal(document.isDirty, true);

		renameModule(workbookPath(), 'VbeDirtyOld', 'VbeDirtyNew');
		await afterTheChangeIsNoticed();

		assert.equal(tabsOn('VbeDirtyOld').length, 1, `the edited editor should stay; tabs now: ${describeTabs()}`);
		assert.equal(tabsOn('VbeDirtyNew').length, 0);
		assert.equal(document.getText().includes("' mine"), true);
	});

	test('does not take a deleted module to another that happens to hold the same code', async () => {
		const same = 'Option Explicit\r\n';
		await agentWrite('VbeTwinA', same);
		await agentWrite('VbeTwinB', same);
		await open(moduleUri('VbeTwinA'));

		deleteModule(workbookPath(), 'VbeTwinA');
		await afterTheChangeIsNoticed();

		assert.equal(tabsOn('VbeTwinB').length, 0, `nothing was renamed; tabs now: ${describeTabs()}`);
	});
});
