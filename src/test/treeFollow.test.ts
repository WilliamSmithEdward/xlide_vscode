// The explorer following the editor, in a real VS Code, under the changes a
// debounce has to survive: tabs switched faster than the tree can reveal, the
// caret crossing procedures, the tree hidden while the editor moves, the
// folder layout, and a project the tree has not loaded yet.
//
// What the view shows is read through xlide.dev.explorerViewState, which asks
// VS Code to redraw and notes whose children it re-reads: the expanded rows,
// and only those.

import * as assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { deleteModule, renameModule } from '../vba/projectService';
import { encodeModuleUri } from '../xlideFileSystem';
import { activate, closeAllEditors, moduleUri, until, workbookPath, workspaceRoot } from './support';

interface ViewState {
	expanded: string[];
	selected: string[];
	activeModule: string | undefined;
}

const viewState = (): Thenable<ViewState> => vscode.commands.executeCommand<ViewState>('xlide.dev.explorerViewState');
const sleep = (ms: number): Promise<void> => new Promise((resolve) => setTimeout(resolve, ms));

/** Long enough for the debounce, the reveal it starts and the caret's row after it. */
const settle = (): Promise<void> => sleep(1500);

async function agentWrite(filePath: string, moduleName: string, source: string): Promise<void> {
	await vscode.lm.invokeTool('xlide_writeModule', {
		input: { filePath, moduleName, source },
		toolInvocationToken: undefined,
	}, new vscode.CancellationTokenSource().token);
}

const MODULES = ['FollowA', 'FollowB', 'FollowC', 'FollowD', 'FollowE'];

function twoProcedures(name: string, folder?: string): string {
	return [
		...(folder ? [`'@Folder("${folder}")`] : []),
		'Option Explicit',
		'',
		`Public Sub ${name}First()`,
		'End Sub',
		'',
		`Public Sub ${name}Second()`,
		'    Dim x As Long',
		'    x = 1',
		'End Sub',
		'',
	].join('\r\n');
}

/** Shows a module with the caret inside its second procedure. */
async function showInSecond(uri: vscode.Uri): Promise<void> {
	const editor = await vscode.window.showTextDocument(uri, { preview: false });
	const line = editor.document.getText().split(/\r?\n/).findIndex((text) => text.includes('Dim x As Long'));
	editor.selection = new vscode.Selection(line, 4, line, 4);
}

const modules = (state: ViewState): string[] => state.expanded.filter((row) => row.startsWith('module:'));

suite('The explorer following the editor', () => {
	suiteSetup(async () => {
		await activate();
		for (const name of MODULES) {
			await agentWrite(workbookPath(), name, twoProcedures(name));
		}
		await vscode.commands.executeCommand('xlide.explorer.focus');
	});

	teardown(async () => {
		await closeAllEditors();
	});

	test('anchor: the probe tells an expanded module from one the accordion folded', async () => {
		await showInSecond(moduleUri('FollowA'));
		await settle();
		const first = await viewState();
		assert.deepEqual(modules(first), ['module:FollowA'], JSON.stringify(first));

		await showInSecond(moduleUri('FollowB'));
		await settle();
		const second = await viewState();
		assert.deepEqual(modules(second), ['module:FollowB'], JSON.stringify(second));
		assert.deepEqual(second.selected, ['sub:Sub FollowBSecond'], JSON.stringify(second));
	});

	test('lands on the last module when tabs switch faster than the tree reveals', async () => {
		// Each module visited once, so its procedure rows exist: the caret's
		// row can then be revealed, which expands the module it sits in.
		for (const name of MODULES) {
			await showInSecond(moduleUri(name));
			await sleep(400);
		}
		for (const gap of [0, 30, 70, 120, 200]) {
			for (const name of [...MODULES, 'FollowB', 'FollowD', 'FollowC']) {
				await showInSecond(moduleUri(name));
				if (gap) {
					await sleep(gap);
				}
			}
			await settle();
			const state = await viewState();
			assert.deepEqual(modules(state), ['module:FollowC'], `switching every ${gap} ms: ${JSON.stringify(state)}`);
			assert.equal(state.activeModule, 'module:FollowC', `switching every ${gap} ms: ${JSON.stringify(state)}`);
			assert.deepEqual(state.selected, ['sub:Sub FollowCSecond'], `switching every ${gap} ms: ${JSON.stringify(state)}`);
		}
	});

	test('marks the procedure the caret ends in, after it crossed several quickly', async () => {
		await showInSecond(moduleUri('FollowD'));
		await settle();
		const editor = vscode.window.activeTextEditor!;
		for (let round = 0; round < 6; round++) {
			for (const line of [3, 7, 3, 7, 3]) {
				editor.selection = new vscode.Selection(line, 1, line, 1);
				await sleep(round * 15);
			}
		}
		await settle();
		const state = await viewState();
		assert.deepEqual(state.selected, ['sub:Sub FollowDFirst'], JSON.stringify(state));
		assert.deepEqual(modules(state), ['module:FollowD'], JSON.stringify(state));
	});

	test('catches up with the editor when the tree is shown again', async () => {
		await showInSecond(moduleUri('FollowA'));
		await settle();
		// Another view in the side bar hides the tree; the editor moves on.
		await vscode.commands.executeCommand('workbench.view.search');
		await sleep(300);
		await showInSecond(moduleUri('FollowE'));
		await settle();
		await vscode.commands.executeCommand('xlide.explorer.focus');
		await settle();
		const state = await viewState();
		assert.deepEqual(modules(state), ['module:FollowE'], JSON.stringify(state));
		assert.deepEqual(state.selected, ['sub:Sub FollowESecond'], JSON.stringify(state));
	});

	test('reveals a module of a project the tree has not loaded', async () => {
		const second = path.join(workspaceRoot(), 'FollowSecond.xlsm');
		fs.copyFileSync(workbookPath(), second);
		try {
			await agentWrite(second, 'SecondOnly', twoProcedures('SecondOnly'));
			await vscode.commands.executeCommand('xlide.refreshExplorer');
			await settle();
			await showInSecond(encodeModuleUri(second, 'SecondOnly'));
			await settle();
			const state = await until(async () => {
				const now = await viewState();
				return now.selected.length > 0 ? now : undefined;
			}, 'the tree should select a row in the other project', 8000).catch(() => viewState());
			assert.ok(state.expanded.includes('project:FollowSecond.xlsm'), JSON.stringify(state));
			assert.deepEqual(modules(state), ['module:SecondOnly'], JSON.stringify(state));
			assert.deepEqual(state.selected, ['sub:Sub SecondOnlySecond'], JSON.stringify(state));
		} finally {
			await closeAllEditors();
			await sleep(500);
			fs.rmSync(second, { force: true });
			await vscode.commands.executeCommand('xlide.refreshExplorer');
		}
	});

	test('keeps the module open and its procedure marked across a save that moves every procedure', async () => {
		await showInSecond(moduleUri('FollowB'));
		await settle();
		const editor = vscode.window.activeTextEditor!;
		// Above every procedure, so the save re-lists them all on new lines.
		await editor.edit((edit) => edit.insert(new vscode.Position(0, 0), "' saved\r\n"));
		await editor.document.save();
		await settle();
		const state = await viewState();
		assert.deepEqual(modules(state), ['module:FollowB'], JSON.stringify(state));
		assert.deepEqual(state.selected, ['sub:Sub FollowBSecond'], JSON.stringify(state));
	});

	test('lands on the last module when refreshes land in the middle of the switching', async () => {
		// An agent writing while you move between tabs: each write refreshes
		// the tree, at whatever point a reveal has reached.
		for (const delay of [0, 20, 50, 100, 200]) {
			for (const name of ['FollowA', 'FollowD', 'FollowB', 'FollowE']) {
				await showInSecond(moduleUri(name));
				await sleep(delay);
				await vscode.commands.executeCommand('xlide.refreshExplorer');
			}
			await settle();
			const state = await viewState();
			assert.deepEqual(modules(state), ['module:FollowE'], `refreshing ${delay} ms after each switch: ${JSON.stringify(state)}`);
			assert.deepEqual(state.selected, ['sub:Sub FollowESecond'], `refreshing ${delay} ms after each switch: ${JSON.stringify(state)}`);
		}
	});

	test('follows the module being edited when the VBE renames it', async () => {
		// The editor moves to the new name before the tree has re-read the
		// project, so the first pass looks for a module the tree does not
		// list yet; the refresh that follows lists it.
		await agentWrite(workbookPath(), 'FollowRenameOld', twoProcedures('FollowRenameOld'));
		await showInSecond(moduleUri('FollowRenameOld'));
		await settle();
		try {
			renameModule(workbookPath(), 'FollowRenameOld', 'FollowRenameNew');
			await until(
				() => (vscode.window.activeTextEditor?.document.uri.toString() === moduleUri('FollowRenameNew').toString() ? true : undefined),
				'the editor should follow the module to its new name',
				8000,
			);
			await settle();
			const state = await viewState();
			assert.deepEqual(modules(state), ['module:FollowRenameNew'], JSON.stringify(state));
			assert.deepEqual(state.selected, ['sub:Sub FollowRenameOldSecond'], JSON.stringify(state));
		} finally {
			await closeAllEditors();
			for (const name of ['FollowRenameOld', 'FollowRenameNew']) {
				try {
					deleteModule(workbookPath(), name);
				} catch {
					// Only one of the two exists, whichever way the test went.
				}
			}
		}
	});

	test('follows the editor that is left after tabs close in a burst', async () => {
		for (const name of MODULES) {
			await showInSecond(moduleUri(name));
		}
		await settle();
		// Ctrl+W held down: E, D and C close, B is in front.
		for (let i = 0; i < 3; i++) {
			await vscode.commands.executeCommand('workbench.action.closeActiveEditor');
			await sleep(20);
		}
		await settle();
		const state = await viewState();
		assert.equal(vscode.window.activeTextEditor?.document.uri.toString(), moduleUri('FollowB').toString());
		assert.deepEqual(modules(state), ['module:FollowB'], JSON.stringify(state));
		assert.deepEqual(state.selected, ['sub:Sub FollowBSecond'], JSON.stringify(state));
	});

	test('keeps one project open when the editor moves between two quickly', async () => {
		const second = path.join(workspaceRoot(), 'FollowOther.xlsm');
		fs.copyFileSync(workbookPath(), second);
		try {
			await agentWrite(second, 'OtherOnly', twoProcedures('OtherOnly'));
			await vscode.commands.executeCommand('xlide.refreshExplorer');
			await settle();
			const other = encodeModuleUri(second, 'OtherOnly');
			for (const gap of [0, 40, 90]) {
				for (const uri of [moduleUri('FollowA'), other, moduleUri('FollowC'), other, moduleUri('FollowE'), other]) {
					await showInSecond(uri);
					if (gap) {
						await sleep(gap);
					}
				}
				await settle();
				const state = await viewState();
				const projects = state.expanded.filter((row) => row.startsWith('project:'));
				assert.deepEqual(projects, ['project:FollowOther.xlsm'], `switching every ${gap} ms: ${JSON.stringify(state)}`);
				assert.deepEqual(modules(state), ['module:OtherOnly'], `switching every ${gap} ms: ${JSON.stringify(state)}`);
				assert.deepEqual(state.selected, ['sub:Sub OtherOnlySecond'], `switching every ${gap} ms: ${JSON.stringify(state)}`);
			}
		} finally {
			await closeAllEditors();
			await sleep(500);
			fs.rmSync(second, { force: true });
			await vscode.commands.executeCommand('xlide.refreshExplorer');
		}
	});

	suite('in the folder layout', () => {
		suiteSetup(async () => {
			await agentWrite(workbookPath(), 'FolderOne', twoProcedures('FolderOne', 'Alpha.Inner'));
			await agentWrite(workbookPath(), 'FolderTwo', twoProcedures('FolderTwo', 'Beta'));
			await agentWrite(workbookPath(), 'FolderThree', twoProcedures('FolderThree', 'Gamma.Deep.Deeper'));
			await vscode.workspace.getConfiguration('xlide').update('explorer.view', 'folders', vscode.ConfigurationTarget.Global);
			await settle();
		});

		suiteTeardown(async () => {
			await vscode.workspace.getConfiguration('xlide').update('explorer.view', undefined, vscode.ConfigurationTarget.Global);
		});

		test('opens only the folders on the way to the last module', async () => {
			const names = ['FolderOne', 'FolderTwo', 'FolderThree'];
			for (const name of names) {
				await showInSecond(moduleUri(name));
				await sleep(400);
			}
			for (const gap of [0, 40, 90]) {
				for (const name of [...names, 'FolderOne', 'FolderThree', 'FolderTwo']) {
					await showInSecond(moduleUri(name));
					if (gap) {
						await sleep(gap);
					}
				}
				await settle();
				const state = await viewState();
				const folders = state.expanded.filter((row) => row.startsWith('folder:'));
				assert.deepEqual(folders, ['folder:Beta'], `switching every ${gap} ms: ${JSON.stringify(state)}`);
				assert.deepEqual(modules(state), ['module:FolderTwo'], `switching every ${gap} ms: ${JSON.stringify(state)}`);
				assert.deepEqual(state.selected, ['sub:Sub FolderTwoSecond'], `switching every ${gap} ms: ${JSON.stringify(state)}`);
			}
		});

		test('follows the module being edited into the folder its annotation now names', async () => {
			await showInSecond(moduleUri('FolderTwo'));
			await settle();
			const editor = vscode.window.activeTextEditor!;
			const annotation = editor.document.getText().split(/\r?\n/).findIndex((text) => text.startsWith("'@Folder("));
			try {
				await editor.edit((edit) => edit.replace(editor.document.lineAt(annotation).range, `'@Folder("Epsilon.Inner")`));
				await settle();
				const state = await viewState();
				const folders = state.expanded.filter((row) => row.startsWith('folder:'));
				assert.deepEqual(folders, ['folder:Epsilon', 'folder:Epsilon.Inner'], JSON.stringify(state));
				assert.deepEqual(modules(state), ['module:FolderTwo'], JSON.stringify(state));
				assert.deepEqual(state.selected, ['sub:Sub FolderTwoSecond'], JSON.stringify(state));

				// The folders it moved into are the ones to fold when the
				// editor moves on, not the one it left.
				await showInSecond(moduleUri('FolderOne'));
				await settle();
				const after = await viewState();
				const open = after.expanded.filter((row) => row.startsWith('folder:'));
				assert.deepEqual(open, ['folder:Alpha', 'folder:Alpha.Inner'], JSON.stringify(after));
			} finally {
				await vscode.window.showTextDocument(moduleUri('FolderTwo'));
				await vscode.commands.executeCommand('workbench.action.files.revert');
			}
		});
	});
});
