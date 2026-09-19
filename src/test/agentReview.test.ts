// Agent review diffs in a real VS Code. An agent often tests with throwaway
// work - a scratch module it creates and deletes - and the review diff of the
// write stayed open after the module was gone, titled for a module that no
// longer existed. Deleting the module now closes its review diffs.
//
// A review diff opens only for a chat-driven write, which a test cannot make,
// so the test opens one the way the review does: the frozen before-image on
// XLIDE's before scheme against the live module document.

import * as assert from 'node:assert/strict';
import * as vscode from 'vscode';
import { activate, closeAllEditors, moduleUri, until, workbookPath } from './support';

async function invoke(tool: string, input: Record<string, unknown>): Promise<void> {
	await vscode.lm.invokeTool(tool, {
		input: { filePath: workbookPath(), ...input },
		toolInvocationToken: undefined,
	}, new vscode.CancellationTokenSource().token);
}

/** Opens a diff the way XLIDE's review does, beside the one already open. */
async function openReviewDiff(moduleName: string): Promise<void> {
	const live = moduleUri(moduleName);
	const before = vscode.Uri.from({ scheme: 'xlide-vba-before', path: live.path, query: 'v1' });
	await vscode.commands.executeCommand('vscode.diff', before, live, `${moduleName}: before agent edit`, { preview: false });
}

function reviewDiffOpen(moduleName: string): boolean {
	const live = moduleUri(moduleName).toString();
	return vscode.window.tabGroups.all.flatMap((group) => group.tabs).some((tab) =>
		tab.input instanceof vscode.TabInputTextDiff && tab.input.modified.toString() === live);
}

suite('Agent review diffs', () => {
	suiteSetup(async () => {
		await activate();
	});

	teardown(async () => {
		await closeAllEditors();
	});

	test('close when the agent deletes the scratch module, and only that module s', async () => {
		await invoke('xlide_writeModule', { moduleName: 'ScratchCheck', source: 'Sub Probe()\r\nEnd Sub\r\n' });
		await invoke('xlide_writeModule', { moduleName: 'KeptWork', source: 'Sub Work()\r\nEnd Sub\r\n' });
		await openReviewDiff('ScratchCheck');
		await openReviewDiff('KeptWork');
		assert.ok(reviewDiffOpen('ScratchCheck') && reviewDiffOpen('KeptWork'), 'both review diffs should be open first');

		await invoke('xlide_deleteModule', { moduleName: 'ScratchCheck' });

		await until(
			() => (reviewDiffOpen('ScratchCheck') ? undefined : true),
			'the deleted module s review diff should close',
		);
		assert.ok(reviewDiffOpen('KeptWork'), 'another module s review diff stays open');
	});
});
