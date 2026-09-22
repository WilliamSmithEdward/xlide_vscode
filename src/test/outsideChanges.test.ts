// Changes made to the workbook outside XLIDE, in a real VS Code: the VBE
// saving it, a git checkout, another window. The suite writes with the engine
// directly, which is exactly that - the extension did not make the write.
// Open modules used to go on showing the old code, and analysis kept using
// it, until the window was reloaded.

import * as assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { readModules } from '../vba/projectService';
import { activate, closeAllEditors, moduleUri, open, until, workbookPath, writeModule as writeOutsideXlide } from './support';

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

suite('Changes made outside XLIDE', () => {
	suiteSetup(async () => {
		await activate();
	});

	teardown(async () => {
		await vscode.commands.executeCommand('workbench.action.files.revert');
		await closeAllEditors();
	});

	test('reach an open module and the modules that call into it', async () => {
		await agentWrite('OutsideCallee', 'Option Explicit\r\n\r\nPublic Sub OldProc()\r\nEnd Sub\r\n');
		await agentWrite('OutsideCaller', 'Option Explicit\r\n\r\nPublic Sub Run()\r\n    OutsideCallee.NewProc\r\nEnd Sub\r\n');
		const callee = await vscode.workspace.openTextDocument(moduleUri('OutsideCallee'));
		const caller = await open(moduleUri('OutsideCaller'));
		const onTheCall = () => vscode.languages.getDiagnostics(caller.uri).filter((d) => d.range.start.line === 3);
		await until(
			() => (onTheCall().length > 0 ? true : undefined),
			'a call to a procedure the module does not have should be reported first',
		);

		const changed = Date.now();
		writeOutsideXlide('OutsideCallee', 'Option Explicit\r\n\r\nPublic Sub OldProc()\r\nEnd Sub\r\n\r\nPublic Sub NewProc()\r\nEnd Sub\r\n');

		await until(
			() => (callee.getText().includes('NewProc') ? true : undefined),
			'the open module should reload with what the workbook now holds',
			8000,
		);
		await until(
			() => (onTheCall().length === 0 ? true : undefined),
			'the calling module should stop reporting NewProc once the workbook has it',
			8000,
		);
		console.log(`    (reloaded and cleared ${Date.now() - changed} ms after the change)`);
		assert.equal(callee.isDirty, false);
	});

	test('let unsaved edits save cleanly when the change did not reach their module', async () => {
		await agentWrite('OutsideEdited', 'Option Explicit\r\n\r\nSub Keep()\r\nEnd Sub\r\n');
		await agentWrite('OutsideOther', 'Option Explicit\r\n');
		const document = await open(moduleUri('OutsideEdited'));
		await vscode.window.activeTextEditor!.edit((edit) => edit.insert(new vscode.Position(4, 0), "' edited\r\n"));
		assert.equal(document.isDirty, true);

		writeOutsideXlide('OutsideOther', 'Option Explicit\r\n\r\nSub Elsewhere()\r\nEnd Sub\r\n');
		await afterTheChangeIsNoticed();

		assert.equal(await document.save(), true, 'a change to another module is no conflict');
		assert.equal(moduleSource('OutsideEdited'), "Option Explicit\r\n\r\nSub Keep()\r\nEnd Sub\r\n' edited\r\n");
		assert.equal(moduleSource('OutsideOther'), 'Option Explicit\r\n\r\nSub Elsewhere()\r\nEnd Sub\r\n');
	});

	test('a save that replaces the file reports a change, which is what the tree follows', async () => {
		// The load-bearing fact behind watchWorkspaceProjectFiles: XLIDE, the
		// Office applications and the MCP server all write by renaming a
		// sibling temp file over the target, and a tree refresh hung on
		// onDidCreate alone never saw one. Measured here rather than assumed,
		// because the whole "the tree does not follow a rename made in the
		// VBE" report turns on which event that write raises.
		const probe = path.join(path.dirname(workbookPath()), 'RenameProbe.xlsm');
		fs.writeFileSync(probe, 'first');
		const watcher = vscode.workspace.createFileSystemWatcher(
			new vscode.RelativePattern(vscode.Uri.file(path.dirname(probe)), 'RenameProbe.xlsm'),
		);
		const events: string[] = [];
		watcher.onDidChange(() => events.push('change'));
		watcher.onDidCreate(() => events.push('create'));
		try {
			// Long enough for the watcher to be listening and for the create
			// above to have been reported and drained.
			await new Promise((resolve) => setTimeout(resolve, 2000));
			events.length = 0;
			const tmp = path.join(path.dirname(probe), '.rename-probe.tmp');
			fs.writeFileSync(tmp, 'second, a different length');
			fs.renameSync(tmp, probe);
			await until(() => (events.length > 0 ? true : undefined), 'the replaced file should be reported', 10000);
			assert.deepEqual([...new Set(events)], ['change']);
		} finally {
			watcher.dispose();
			fs.rmSync(probe, { force: true });
		}
	});

	test('refuse to save unsaved edits over a change that reached their module', async () => {
		await agentWrite('OutsideContested', 'Option Explicit\r\n\r\nSub Mine()\r\nEnd Sub\r\n');
		const document = await open(moduleUri('OutsideContested'));
		await vscode.window.activeTextEditor!.edit((edit) => edit.insert(new vscode.Position(4, 0), "' mine\r\n"));

		writeOutsideXlide('OutsideContested', 'Option Explicit\r\n\r\nSub Theirs()\r\nEnd Sub\r\n');
		await afterTheChangeIsNoticed();

		assert.equal(await document.save(), false, 'the save should stop at the conflict');
		assert.equal(moduleSource('OutsideContested'), 'Option Explicit\r\n\r\nSub Theirs()\r\nEnd Sub\r\n');
	});
});
