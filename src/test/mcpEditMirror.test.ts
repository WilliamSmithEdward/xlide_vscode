// The MCP server's edits mirrored into a real VS Code. The suite does what the
// server does: it writes the workbook with the engine directly, in its own
// right rather than through XLIDE, then reports the write to the loopback API
// the window serves, found through the record the window wrote.
//
// A review diff opens for these, unlike for a tool call a test makes, because
// the report is what presents it.

import * as assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { deleteModule, readModules, renameModule, writeModule as engineWrite } from '../vba/projectService';
import { xlideApiRecordName, xlideApiStateDir } from '../xlideApiServer';
import { activate, closeAllEditors, moduleUri, until, workbookPath } from './support';

interface Record {
	port: number;
	token: string;
	product: string;
	protocol: number;
	workspaceFolders: string[];
}

interface ViewState {
	expanded: string[];
	selected: string[];
	activeModule: string | undefined;
}

const sleep = (ms: number): Promise<void> => new Promise((resolve) => setTimeout(resolve, ms));
const viewState = (): Thenable<ViewState> => vscode.commands.executeCommand<ViewState>('xlide.dev.explorerViewState');
const modules = (state: ViewState): string[] => state.expanded.filter((row) => row.startsWith('module:'));

/** Long enough for the refresh a report causes, and the tree following the editor after it. */
const settle = (): Promise<void> => sleep(1500);

/** The record this window wrote: the extension host is this process. */
async function record(): Promise<Record> {
	const file = path.join(xlideApiStateDir(), xlideApiRecordName(process.pid));
	return until(() => (fs.existsSync(file) ? JSON.parse(fs.readFileSync(file, 'utf8')) as Record : undefined),
		`the window should write ${file}`);
}

async function report(route: string, body: object): Promise<{ status: number; json: { shown?: boolean; review?: string } }> {
	const { port, token } = await record();
	const response = await fetch(`http://127.0.0.1:${port}/${token}/${route}`, {
		method: 'POST',
		headers: { 'Content-Type': 'application/json' },
		body: JSON.stringify(body),
	});
	return { status: response.status, json: await response.json() as { shown?: boolean; review?: string } };
}

function source(moduleName: string): string | undefined {
	return readModules(workbookPath()).find((module) => module.name === moduleName)?.source;
}

/** What the server does for one tool call: write the module itself, then report it. */
async function serverWrites(moduleName: string, code: string): Promise<{ status: number; json: { shown?: boolean; review?: string } }> {
	const before = source(moduleName);
	engineWrite(workbookPath(), moduleName, code, 'standard');
	return report('agent-edit', {
		file: workbookPath(),
		module: moduleName,
		// The server's header-stripped body, which can keep blank lines the
		// engine drops: the review must still see the module as unchanged there.
		before: before === undefined ? null : `\r\n${before}`,
		beforeExisted: before !== undefined,
		after: code,
		afterExists: true,
		kind: 'write',
	});
}

function reviewDiffOpen(moduleName: string): boolean {
	const live = moduleUri(moduleName).toString();
	return vscode.window.tabGroups.all.flatMap((group) => group.tabs).some((tab) =>
		tab.input instanceof vscode.TabInputTextDiff
		&& tab.input.original.scheme === 'xlide-vba-before'
		&& tab.input.modified.toString() === live);
}

function procedures(name: string, extra = ''): string {
	return [
		'Option Explicit',
		'',
		`Public Sub ${name}First()`,
		'End Sub',
		'',
		`Public Sub ${name}Second()`,
		'    Dim x As Long',
		`    x = 1${extra}`,
		'End Sub',
		'',
	].join('\r\n');
}

suite('The MCP server\'s edits, mirrored', () => {
	suiteSetup(async () => {
		await activate();
	});

	teardown(async () => {
		await closeAllEditors();
	});

	test('the window says where it listens, and answers hello', async () => {
		const found = await record();
		assert.equal(found.product, 'xlide_vscode');
		assert.equal(found.protocol, 1);
		assert.deepEqual(found.workspaceFolders, vscode.workspace.workspaceFolders?.map((folder) => folder.uri.fsPath));

		const hello = await fetch(`http://127.0.0.1:${found.port}/${found.token}/hello`);
		assert.deepEqual(await hello.json(), { product: 'xlide_vscode', protocol: 1 });
	});

	test('a module the server wrote gets a diff, and Revert puts back what it had', async () => {
		engineWrite(workbookPath(), 'MirrorEdited', procedures('MirrorEdited'), 'standard');
		const original = source('MirrorEdited');

		const answer = await serverWrites('MirrorEdited', procedures('MirrorEdited', ' + 1'));

		assert.deepEqual(answer, { status: 200, json: { shown: true, review: 'pending' } });
		await until(() => (reviewDiffOpen('MirrorEdited') ? true : undefined), 'the review diff should open');
		await vscode.commands.executeCommand('xlide.revertAgentChange', { filePath: workbookPath(), moduleName: 'MirrorEdited' });
		assert.equal(source('MirrorEdited'), original);
		await until(() => (reviewDiffOpen('MirrorEdited') ? undefined : true), 'the diff should close once the change is reverted');
	});

	test('a write that changed nothing leaves nothing to review', async () => {
		engineWrite(workbookPath(), 'MirrorSame', procedures('MirrorSame'), 'standard');

		const answer = await serverWrites('MirrorSame', procedures('MirrorSame'));

		assert.deepEqual(answer.json, { shown: true, review: 'none' });
		await sleep(500);
		assert.equal(reviewDiffOpen('MirrorSame'), false);
	});

	test('a module the server deleted takes its review, and its diff, with it', async () => {
		engineWrite(workbookPath(), 'MirrorDeleted', procedures('MirrorDeleted'), 'standard');
		await serverWrites('MirrorDeleted', procedures('MirrorDeleted', ' + 2'));
		await until(() => (reviewDiffOpen('MirrorDeleted') ? true : undefined), 'the review diff should open');

		deleteModule(workbookPath(), 'MirrorDeleted');
		const answer = await report('agent-edit', {
			file: workbookPath(),
			module: 'MirrorDeleted',
			before: procedures('MirrorDeleted', ' + 2'),
			beforeExisted: true,
			after: null,
			afterExists: false,
			kind: 'delete',
		});

		assert.deepEqual(answer.json, { shown: true, review: 'none' });
		await until(() => (reviewDiffOpen('MirrorDeleted') ? undefined : true), 'the deleted module s diff should close');
	});

	test('a module the server renamed keeps its review under the new name', async () => {
		engineWrite(workbookPath(), 'MirrorOld', procedures('MirrorOld'), 'standard');
		await serverWrites('MirrorOld', procedures('MirrorOld', ' + 3'));
		await until(() => (reviewDiffOpen('MirrorOld') ? true : undefined), 'the review diff should open');

		renameModule(workbookPath(), 'MirrorOld', 'MirrorNew');
		const answer = await report('module-renamed', { file: workbookPath(), from: 'MirrorOld', to: 'MirrorNew' });

		assert.deepEqual(answer.json, { shown: true, review: 'pending' });
		await until(() => (reviewDiffOpen('MirrorNew') && !reviewDiffOpen('MirrorOld') ? true : undefined),
			'the diff should follow the module to its new name');
		await vscode.commands.executeCommand('xlide.keepAgentChange', { filePath: workbookPath(), moduleName: 'MirrorNew' });
	});

	test('a file this window does not show is left alone', async () => {
		const elsewhere = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-mirror-')), 'Elsewhere.xlsm');
		fs.copyFileSync(workbookPath(), elsewhere);
		try {
			const answer = await report('file-changed', { file: elsewhere, what: 'cells' });
			assert.deepEqual(answer, { status: 200, json: { shown: false, review: 'none' } });
		} finally {
			fs.rmSync(path.dirname(elsewhere), { recursive: true, force: true });
		}
	});

	test('a caller without the token gets nothing', async () => {
		const { port } = await record();
		const response = await fetch(`http://127.0.0.1:${port}/not-the-token/hello`);
		assert.equal(response.status, 404);
	});

	suite('and the tree following the editor', () => {
		/** Shows a module with the caret inside its second procedure. */
		async function showInSecond(moduleName: string): Promise<void> {
			const editor = await vscode.window.showTextDocument(moduleUri(moduleName), { preview: false });
			const line = editor.document.getText().split(/\r?\n/).findIndex((text) => text.includes('Dim x As Long'));
			editor.selection = new vscode.Selection(line, 4, line, 4);
		}

		async function assertStillOn(moduleName: string, after: string): Promise<void> {
			await settle();
			const state = await viewState();
			assert.deepEqual(modules(state), [`module:${moduleName}`], `${after}: ${JSON.stringify(state)}`);
			assert.deepEqual(state.selected, [`sub:Sub ${moduleName}Second`], `${after}: ${JSON.stringify(state)}`);
		}

		suiteSetup(async () => {
			engineWrite(workbookPath(), 'MirrorEditing', procedures('MirrorEditing'), 'standard');
			engineWrite(workbookPath(), 'MirrorElsewhere', procedures('MirrorElsewhere'), 'standard');
			await vscode.commands.executeCommand('xlide.refreshExplorer');
			await vscode.commands.executeCommand('xlide.explorer.focus');
		});

		test('keeps the module being edited open, its procedure marked, as the server writes the file', async () => {
			// Review Agent Writes off: the reloads alone, with no diff tab
			// taking the editor's place.
			const config = vscode.workspace.getConfiguration('xlide');
			await config.update('agent.showWriteDiffs', false, vscode.ConfigurationTarget.Global);
			try {
				await showInSecond('MirrorEditing');
				await assertStillOn('MirrorEditing', 'before any report');

				await serverWrites('MirrorElsewhere', procedures('MirrorElsewhere', ' + 4'));
				await assertStillOn('MirrorEditing', 'after the server wrote another module');

				await serverWrites('MirrorEditing', procedures('MirrorEditing', ' + 5'));
				await until(() => (vscode.window.activeTextEditor?.document.getText().includes('x = 1 + 5') ? true : undefined),
					'the open module should reload with the server s write');
				await assertStillOn('MirrorEditing', 'after the server wrote the module being edited');

				await report('file-changed', { file: workbookPath(), what: 'cells' });
				await assertStillOn('MirrorEditing', 'after the server changed cells');
			} finally {
				await config.update('agent.showWriteDiffs', undefined, vscode.ConfigurationTarget.Global);
			}
		});

		test('opens the module in the tab in front when a review diff takes the editor\'s place', async () => {
			await showInSecond('MirrorEditing');
			await settle();

			await serverWrites('MirrorElsewhere', procedures('MirrorElsewhere', ' + 6'));
			await until(() => (reviewDiffOpen('MirrorElsewhere') ? true : undefined), 'the review diff should open');
			await settle();

			const front = vscode.window.activeTextEditor?.document.uri.path ?? '';
			const inFront = path.posix.basename(front).replace(/\.bas$/, '');
			const state = await viewState();
			assert.deepEqual(modules(state), [`module:${inFront}`], JSON.stringify({ front, state }));
			await vscode.commands.executeCommand('xlide.keepAgentChange', { filePath: workbookPath(), moduleName: 'MirrorElsewhere' });
		});
	});
});
