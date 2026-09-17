// Compare with git in a real VS Code: the workbook is committed, a module is
// changed through XLIDE, and the command opens a diff whose left side is the
// module as committed.

import * as assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import * as vscode from 'vscode';
import { readModules } from '../vba/projectService';
import { activate, closeAllEditors, moduleUri, until, workbookPath, workspaceRoot, writeModule } from './support';

function git(...args: string[]): string {
	return execFileSync('git', args, { cwd: workspaceRoot(), stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8').trim();
}

function gitAvailable(): boolean {
	try {
		execFileSync('git', ['--version'], { stdio: 'ignore' });
		return true;
	} catch {
		return false;
	}
}

function openDiffTab(): vscode.TabInputTextDiff | undefined {
	for (const group of vscode.window.tabGroups.all) {
		for (const tab of group.tabs) {
			if (tab.input instanceof vscode.TabInputTextDiff && tab.input.original.scheme === 'xlide-vba-git') {
				return tab.input;
			}
		}
	}
	return undefined;
}

suite('Compare with git', () => {
	suiteSetup(async function () {
		if (!gitAvailable()) {
			this.skip();
		}
		await activate();
		await writeModule('GitProbe', 'Option Explicit\r\n\r\nSub Before()\r\n    Debug.Print "committed"\r\nEnd Sub\r\n');
		git('init', '-q');
		git('config', 'user.email', 'test@example.com');
		git('config', 'user.name', 'Test');
		git('add', 'FormFixture.xlsm');
		git('commit', '-q', '-m', 'commit the workbook');
		await writeModule('GitProbe', 'Option Explicit\r\n\r\nSub After()\r\n    Debug.Print "changed"\r\nEnd Sub\r\n');
	});

	teardown(async () => {
		await closeAllEditors();
	});

	test('Compare Module with Git HEAD opens the committed module against the live one', async () => {
		await vscode.commands.executeCommand('xlide.compareModuleWithHead', { filePath: workbookPath(), moduleName: 'GitProbe' });
		const diff = await until(openDiffTab, 'a diff against the committed module should open');
		assert.equal(diff.modified.toString(), moduleUri('GitProbe').toString());
		const committed = await vscode.workspace.openTextDocument(diff.original);
		assert.ok(committed.getText().includes('Sub Before()'), committed.getText());
		assert.ok(!committed.getText().includes('Sub After()'), committed.getText());
		const live = await vscode.workspace.openTextDocument(diff.modified);
		assert.ok(live.getText().includes('Sub After()'), live.getText());
	});

	test('Restore Module from Git HEAD writes the committed module back into the workbook', async () => {
		await vscode.commands.executeCommand('xlide.restoreModuleFromHead', { filePath: workbookPath(), moduleName: 'GitProbe' }, { withoutPrompt: true });
		const restored = await until(() => {
			const source = readModules(workbookPath()).find((module) => module.name === 'GitProbe')?.source ?? '';
			return source.includes('Sub Before()') && !source.includes('Sub After()') ? source : undefined;
		}, 'GitProbe should be back to its committed text on disk');
		assert.ok(restored.includes('"committed"'), restored);
		// Put the change back so the workbook is left the way the suite set it up.
		await writeModule('GitProbe', 'Option Explicit\r\n\r\nSub After()\r\n    Debug.Print "changed"\r\nEnd Sub\r\n');
	});

	test('the xlide_gitChanges tool hands an agent the module diffs against HEAD', async () => {
		const result = await vscode.lm.invokeTool('xlide_gitChanges', { input: { filePath: workbookPath() }, toolInvocationToken: undefined });
		const text = result.content
			.map((part) => (part instanceof vscode.LanguageModelTextPart ? part.value : ''))
			.join('');
		const report = JSON.parse(text) as { tracked: boolean; changes: Array<{ name: string; kind: string; diff: string }> };
		assert.equal(report.tracked, true);
		const probe = report.changes.find((change) => change.name === 'GitProbe');
		assert.ok(probe, text);
		assert.equal(probe.kind, 'modified');
		assert.ok(probe.diff.includes('-Sub Before()') && probe.diff.includes('+Sub After()'), probe.diff);
	});

	test('Show Module History lists the commits that changed the module and diffs one against the one before', async () => {
		git('commit', '-q', '-a', '-m', 'second change');
		const head = git('rev-parse', 'HEAD');
		await vscode.commands.executeCommand('xlide.moduleHistory', { filePath: workbookPath(), moduleName: 'GitProbe' }, { commit: head });
		const diff = await until(openDiffTab, 'a diff between the two committed versions should open');
		assert.equal(diff.modified.scheme, 'xlide-vba-git');
		const before = await vscode.workspace.openTextDocument(diff.original);
		const after = await vscode.workspace.openTextDocument(diff.modified);
		assert.ok(before.getText().includes('Sub Before()'), before.getText());
		assert.ok(after.getText().includes('Sub After()'), after.getText());
	});
});
