// Canonical casing in a real VS Code: XLIDE recases text the user types, and
// nothing else. When an agent writes a module that is open, VS Code reloads
// the document from the workbook, and the reload arrives as an ordinary
// content change. It was taken for typing: the agent's `debug.print 1` came
// back as `Debug.Print 1` and the document was left dirty, out of step with
// the workbook, holding edits nobody made.

import * as assert from 'node:assert/strict';
import * as vscode from 'vscode';
import { activate, closeAllEditors, moduleUri, open, until, workbookPath } from './support';

/** Writes a module the way an agent does: through XLIDE's write tool. */
async function agentWrite(moduleName: string, source: string): Promise<void> {
	await vscode.lm.invokeTool('xlide_writeModule', {
		input: { filePath: workbookPath(), moduleName, source },
		toolInvocationToken: undefined,
	}, new vscode.CancellationTokenSource().token);
}

/** Longer than the casing controller's idle pause, so a recase would have run. */
function pastTheIdlePause(): Promise<void> {
	return new Promise((resolve) => setTimeout(resolve, 800));
}

suite('Canonical casing', () => {
	suiteSetup(async () => {
		await activate();
	});

	teardown(async () => {
		await closeAllEditors();
	});

	test('leaves an agent s write to the open module as the agent wrote it', async () => {
		await agentWrite('CaseReload', 'option explicit\r\n\r\nsub a()\r\nend sub\r\n');
		const document = await open(moduleUri('CaseReload'));
		const rewritten = 'option explicit\r\n\r\nsub a()\r\n    debug.print 1\r\nend sub\r\n';

		await agentWrite('CaseReload', rewritten);
		await until(
			() => (document.getText() === rewritten ? true : undefined),
			'the open module should reload with what the agent wrote',
		);
		await pastTheIdlePause();

		assert.equal(document.getText(), rewritten);
		assert.equal(document.isDirty, false, 'the module should still match the workbook');
	});

	test('recases what the user types under a short auto-save delay', async () => {
		// The save came before the pause, the pause found the module matching
		// its file, and `dim x as long` was left as `Dim x As long`.
		const files = vscode.workspace.getConfiguration('files');
		await files.update('autoSave', 'afterDelay', vscode.ConfigurationTarget.Workspace);
		await files.update('autoSaveDelay', 100, vscode.ConfigurationTarget.Workspace);
		try {
			await agentWrite('CaseAutoSave', 'Option Explicit\r\n\r\nSub T()\r\n\r\nEnd Sub\r\n');
			const document = await open(moduleUri('CaseAutoSave'));
			vscode.window.activeTextEditor!.selection = new vscode.Selection(3, 0, 3, 0);
			for (const ch of 'dim x as long') {
				await vscode.commands.executeCommand('type', { text: ch });
				await new Promise((resolve) => setTimeout(resolve, 30));
			}
			await until(
				() => (document.lineAt(3).text === 'Dim x As Long' && !document.isDirty ? true : undefined),
				'the typed line should be recased and saved',
				5000,
			);
		} finally {
			await files.update('autoSave', undefined, vscode.ConfigurationTarget.Workspace);
			await files.update('autoSaveDelay', undefined, vscode.ConfigurationTarget.Workspace);
		}
	});

	test('changes nothing when a module is only opened and moved through', async () => {
		const source = 'option explicit\r\n\r\nsub b()\r\ndim x as long\r\nend sub\r\n';
		await agentWrite('CaseOpen', source);
		const document = await open(moduleUri('CaseOpen'));
		const editor = vscode.window.activeTextEditor!;
		for (const line of [1, 2, 3, 4, 0, 3]) {
			editor.selection = new vscode.Selection(line, 0, line, 0);
			await new Promise((resolve) => setTimeout(resolve, 100));
		}
		await pastTheIdlePause();

		assert.equal(document.getText(), source);
		assert.equal(document.isDirty, false);
	});
});
