// The git compare commands: what they open, what they say when git cannot
// help, and how a project's changed modules are listed.

import { beforeEach, describe, expect, it, vi } from 'vitest';

const ui = vi.hoisted(() => ({
	quickPick: vi.fn(async (items: unknown[]) => items[0]),
	status: vi.fn(),
	openTextDocument: vi.fn(),
	applyEdit: vi.fn(async () => true),
	showTextDocument: vi.fn(async () => undefined),
	showWarningMessage: vi.fn(async () => 'Restore'),
	withProgress: vi.fn((_options: unknown, task: (progress: { report(): void }, token: { isCancellationRequested: boolean }) => unknown) =>
		task({ report: () => undefined }, { isCancellationRequested: false })),
}));

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
	EndOfLine: { LF: 1, CRLF: 2 },
	ProgressLocation: { Notification: 15 },
	WorkspaceEdit: class {
		replacements: Array<{ uri: unknown; newText: string }> = [];
		replace(uri: unknown, _range: unknown, newText: string): void { this.replacements.push({ uri, newText }); }
	},
	window: {
		showQuickPick: ui.quickPick,
		setStatusBarMessage: ui.status,
		showTextDocument: ui.showTextDocument,
		showWarningMessage: ui.showWarningMessage,
		withProgress: ui.withProgress,
		activeTextEditor: undefined,
	},
	workspace: {
		openTextDocument: ui.openTextDocument,
		applyEdit: ui.applyEdit,
	},
}));

import * as vscode from 'vscode';
import type { GitRunResult } from '../src/gitFileHistory';
import {
	XLIDE_GIT_SCHEME,
	gitChangesReport,
	moduleChanges,
	moduleHistoryEntries,
	registerGitCompareCommands,
	resetGitSnapshotCacheForTests,
	type GitModuleCompareDeps,
	type ModuleSnapshot,
} from '../src/gitModuleCompare';

const PROJECT = process.platform === 'win32' ? 'C:\\work\\Book.xlsm' : '/work/Book.xlsm';
const ROOT = process.platform === 'win32' ? 'C:\\work' : '/work';

function result(code: number, stdout = ''): GitRunResult {
	return { code, stdout: Buffer.from(stdout), stderr: '' };
}

/** A git that knows one repository at ROOT with the workbook tracked. */
function fakeGit(overrides: { tracked?: boolean; inRepo?: boolean } = {}) {
	return {
		run: vi.fn(async (args: readonly string[]) => {
			if (args[0] === 'rev-parse') {
				return overrides.inRepo === false ? result(128) : result(0, `${ROOT.replace(/\\/g, '/')}\n`);
			}
			if (args[0] === 'ls-files') {
				return result(overrides.tracked === false ? 1 : 0);
			}
			if (args[0] === 'log') {
				return result(0, 'aaaa1111\u001faaaa111\u001fAnn\u001f2026-09-01\u001fsecond\nbbbb2222\u001fbbbb222\u001fBob\u001f2026-08-01\u001ffirst\n');
			}
			return result(0);
		}),
	};
}

function deps(
	committed: ModuleSnapshot[] | undefined,
	current: ModuleSnapshot[],
	git = fakeGit(),
): GitModuleCompareDeps & { git: ReturnType<typeof fakeGit> } {
	return {
		git,
		currentModules: vi.fn(async () => current),
		modulesAtRevision: vi.fn(async () => committed),
	};
}

/** Registers the commands against `d` and hands back THIS registration's handlers. */
function handlers(d: GitModuleCompareDeps) {
	const commandsBefore = vi.mocked(vscode.commands.registerCommand).mock.calls.length;
	const providersBefore = vi.mocked(vscode.workspace.registerTextDocumentContentProvider).mock.calls.length;
	const disposables = registerGitCompareCommands({} as never, d);
	const registered = vi.mocked(vscode.commands.registerCommand).mock.calls.slice(commandsBefore);
	const handler = (id: string) => {
		const call = registered.find(([name]) => name === id);
		if (!call) { throw new Error(`command ${id} not registered`); }
		return call[1] as (node?: unknown) => Promise<void>;
	};
	const provider = vi.mocked(vscode.workspace.registerTextDocumentContentProvider).mock.calls
		.slice(providersBefore)
		.find(([scheme]) => scheme === XLIDE_GIT_SCHEME)?.[1] as { provideTextDocumentContent(uri: unknown): string };
	return { handler, provider, disposables };
}

function setActiveEditor(editor: unknown): void {
	(vscode.window as { activeTextEditor?: unknown }).activeTextEditor = editor;
}

function diffCalls() {
	return vi.mocked(vscode.commands.executeCommand).mock.calls
		.filter(([command]) => command === 'vscode.diff')
		.map(([, left, right, title]) => ({
			left: left as { scheme: string; path: string; toString(): string },
			right: right as { scheme: string; path: string; toString(): string },
			title: title as string,
		}));
}

beforeEach(() => {
	vi.mocked(vscode.commands.registerCommand).mockClear();
	vi.mocked(vscode.commands.executeCommand).mockClear();
	vi.mocked(vscode.workspace.registerTextDocumentContentProvider).mockClear();
	vi.mocked(vscode.window.showInformationMessage).mockClear();
	ui.showWarningMessage.mockClear();
	ui.quickPick.mockClear();
	ui.status.mockClear();
	ui.openTextDocument.mockReset();
	ui.applyEdit.mockClear();
	ui.showTextDocument.mockClear();
	ui.withProgress.mockClear();
	resetGitSnapshotCacheForTests();
	setActiveEditor(undefined);
});

/** A module document the way the restore command reads and edits it. */
function fakeDocument(text: string, dirty = false, eol: 1 | 2 = 2) {
	return {
		uri: { scheme: 'xlide-vba', path: '/C:/work/Book.xlsm/Module1.bas', toString: () => 'xlide-vba:/C:/work/Book.xlsm/Module1.bas' },
		eol,
		isDirty: dirty,
		getText: () => text,
		positionAt: (offset: number) => ({ line: 0, character: offset }),
		save: vi.fn(async () => true),
	};
}

describe('Compare Module with Git HEAD', () => {
	it('opens a diff with the committed text on the left and the live module on the right', async () => {
		const d = deps([{ name: 'Module1', source: 'old' }], [{ name: 'Module1', source: 'new' }]);
		const { handler, provider } = handlers(d);
		await handler('xlide.compareModuleWithHead')({ filePath: PROJECT, moduleName: 'Module1' });

		const [diff] = diffCalls();
		expect(diff).toBeDefined();
		expect(diff.left.scheme).toBe(XLIDE_GIT_SCHEME);
		expect(diff.left.path.endsWith('/Module1.bas')).toBe(true);
		expect(diff.right.scheme).toBe('xlide-vba');
		expect(diff.title).toBe('Module1: HEAD \u2194 current');
		expect(provider.provideTextDocumentContent(diff.left)).toBe('old');
		expect(d.modulesAtRevision).toHaveBeenCalledWith(PROJECT, expect.objectContaining({ relativePath: 'Book.xlsm' }), 'HEAD');
	});

	it('takes the module from the active editor when no tree row was given', async () => {
		const d = deps([{ name: 'Sheet1', source: 'a' }], [{ name: 'Sheet1', source: 'b' }]);
		const { handler } = handlers(d);
		setActiveEditor({ document: { uri: vscode.Uri.from({ scheme: 'xlide-vba', path: `/${PROJECT.replace(/\\/g, '/')}/Sheet1.bas` }) } });
		await handler('xlide.compareModuleWithHead')();
		expect(diffCalls()[0]?.title).toBe('Sheet1: HEAD \u2194 current');
	});

	it('diffs a module the commit lacks against nothing, and says so in the title', async () => {
		const d = deps([], [{ name: 'NewOne', source: 'x' }]);
		const { handler, provider } = handlers(d);
		await handler('xlide.compareModuleWithHead')({ filePath: PROJECT, moduleName: 'NewOne' });
		const [diff] = diffCalls();
		expect(diff.title).toBe('NewOne: not in HEAD \u2194 current');
		expect(provider.provideTextDocumentContent(diff.left)).toBe('');
	});

	it('explains a file outside a repository, an untracked file, and a file absent from the revision', async () => {
		const outside = deps([], [], fakeGit({ inRepo: false }));
		await handlers(outside).handler('xlide.compareModuleWithHead')({ filePath: PROJECT, moduleName: 'M' });
		expect(vi.mocked(vscode.window.showInformationMessage).mock.calls[0][0]).toContain('not inside a git repository');

		const untracked = deps([], [], fakeGit({ tracked: false }));
		await handlers(untracked).handler('xlide.compareModuleWithHead')({ filePath: PROJECT, moduleName: 'M' });
		expect(vi.mocked(vscode.window.showInformationMessage).mock.calls[1][0]).toContain('not tracked by git');

		const absent = deps(undefined, []);
		await handlers(absent).handler('xlide.compareModuleWithHead')({ filePath: PROJECT, moduleName: 'M' });
		expect(vi.mocked(vscode.window.showInformationMessage).mock.calls[2][0]).toContain('is not in HEAD');
		expect(diffCalls()).toHaveLength(0);
	});

	it('warns when git itself cannot run', async () => {
		const { GitUnavailableError } = await import('../src/gitFileHistory');
		const broken = deps([], [], { run: vi.fn(async () => { throw new GitUnavailableError('git', 'ENOENT'); }) } as never);
		await handlers(broken).handler('xlide.compareModuleWithHead')({ filePath: PROJECT, moduleName: 'M' });
		expect(vi.mocked(vscode.window.showWarningMessage).mock.calls[0][0]).toContain('git.path');
	});
});

describe('Compare Module with Git Revision', () => {
	it('offers the commits that touched the file and diffs against the one picked', async () => {
		const d = deps([{ name: 'Module1', source: 'then' }], [{ name: 'Module1', source: 'now' }]);
		const { handler, provider } = handlers(d);
		await handler('xlide.compareModuleWithRevision')({ filePath: PROJECT, moduleName: 'Module1' });

		const offered = ui.quickPick.mock.calls[0][0] as Array<{ label: string; description: string }>;
		expect(offered.map((item) => item.label)).toEqual(['$(git-commit) aaaa111  second', '$(git-commit) bbbb222  first']);
		expect(offered[0].description).toBe('Ann, 2026-09-01');
		expect(d.modulesAtRevision).toHaveBeenCalledWith(PROJECT, expect.anything(), 'aaaa1111');
		const [diff] = diffCalls();
		expect(diff.title).toBe('Module1: aaaa111 \u2194 current');
		expect(provider.provideTextDocumentContent(diff.left)).toBe('then');
	});
});

describe('Compare File with Git HEAD', () => {
	it('lists modified, added and removed modules and opens the diff of the pick', async () => {
		const d = deps(
			[{ name: 'Same', source: 'x' }, { name: 'Changed', source: 'a' }, { name: 'Gone', source: 'g' }],
			[{ name: 'Same', source: 'x\r\n' }, { name: 'Changed', source: 'b' }, { name: 'Fresh', source: 'f' }],
		);
		ui.quickPick.mockImplementationOnce(async (items: unknown[]) => (items as Array<{ label: string }>).find((item) => item.label.includes('Gone')));
		const { handler, provider } = handlers(d);
		await handler('xlide.compareProjectWithHead')({ filePath: PROJECT });

		const offered = ui.quickPick.mock.calls[0][0] as Array<{ label: string; description: string }>;
		expect(offered.map((item) => `${item.label} ${item.description}`)).toEqual([
			'$(diff-modified) Changed modified',
			'$(diff-added) Fresh added',
			'$(diff-removed) Gone removed',
		]);
		const [diff] = diffCalls();
		expect(diff.title).toBe('Gone: HEAD \u2194 removed');
		expect(provider.provideTextDocumentContent(diff.left)).toBe('g');
		expect(diff.right.scheme).toBe(XLIDE_GIT_SCHEME);
		expect(provider.provideTextDocumentContent(diff.right)).toBe('');
	});

	it('says so in the status bar when nothing differs', async () => {
		const d = deps([{ name: 'Same', source: 'x' }], [{ name: 'same', source: 'x' }]);
		await handlers(d).handler('xlide.compareProjectWithHead')({ filePath: PROJECT });
		expect(ui.status.mock.calls[0][0]).toContain('No VBA changes');
		expect(ui.quickPick).not.toHaveBeenCalled();
	});
});

describe('Restore Module from Git HEAD', () => {
	it('replaces the module text with the committed text, in the document\'s line endings, and saves', async () => {
		const document = fakeDocument('Sub After()\r\nEnd Sub\r\n');
		ui.openTextDocument.mockResolvedValue(document);
		const d = deps([{ name: 'Module1', source: 'Sub Before()\nEnd Sub\n' }], [{ name: 'Module1', source: 'Sub After()\r\nEnd Sub\r\n' }]);
		const { handler } = handlers(d);

		await handler('xlide.restoreModuleFromHead')({ filePath: PROJECT, moduleName: 'Module1' });

		expect(ui.showWarningMessage.mock.calls[0][0]).toBe('Restore "Module1" from HEAD?');
		const edit = ui.applyEdit.mock.calls[0][0] as { replacements: Array<{ newText: string }> };
		expect(edit.replacements[0].newText).toBe('Sub Before()\r\nEnd Sub\r\n');
		expect(ui.showTextDocument).toHaveBeenCalledWith(document, { preview: true });
		expect(document.save).toHaveBeenCalledTimes(1);
		expect(ui.status.mock.calls[0][0]).toBe('XLIDE: Restored Module1 from HEAD');
	});

	it('leaves a document with unsaved edits unsaved after restoring into it', async () => {
		const document = fakeDocument('Sub Typing()\r\n', true);
		ui.openTextDocument.mockResolvedValue(document);
		const d = deps([{ name: 'Module1', source: 'Sub Before()\r\nEnd Sub\r\n' }], [{ name: 'Module1', source: 'Sub After()\r\nEnd Sub\r\n' }]);
		await handlers(d).handler('xlide.restoreModuleFromHead')({ filePath: PROJECT, moduleName: 'Module1' }, { withoutPrompt: true });

		expect(ui.showWarningMessage).not.toHaveBeenCalled();
		expect(ui.applyEdit).toHaveBeenCalledTimes(1);
		expect(document.save).not.toHaveBeenCalled();
		expect(ui.status.mock.calls[0][0]).toContain('save to write it');
	});

	it('does nothing when the module already matches HEAD, or when the user declines', async () => {
		ui.openTextDocument.mockResolvedValue(fakeDocument('Sub Same()\r\nEnd Sub\r\n'));
		const same = deps([{ name: 'Module1', source: 'Sub Same()\nEnd Sub\n' }], [{ name: 'Module1', source: 'Sub Same()\r\nEnd Sub\r\n' }]);
		await handlers(same).handler('xlide.restoreModuleFromHead')({ filePath: PROJECT, moduleName: 'Module1' });
		expect(ui.status.mock.calls[0][0]).toBe('XLIDE: Module1 already matches HEAD');
		expect(ui.applyEdit).not.toHaveBeenCalled();

		ui.openTextDocument.mockResolvedValue(fakeDocument('Sub After()\r\nEnd Sub\r\n'));
		ui.showWarningMessage.mockResolvedValueOnce(undefined as never);
		const declined = deps([{ name: 'Module1', source: 'Sub Before()\r\nEnd Sub\r\n' }], [{ name: 'Module1', source: 'Sub After()\r\nEnd Sub\r\n' }]);
		await handlers(declined).handler('xlide.restoreModuleFromHead')({ filePath: PROJECT, moduleName: 'Module1' });
		expect(ui.applyEdit).not.toHaveBeenCalled();
	});

	it('refuses a module the last commit does not have, and says why', async () => {
		const d = deps([], [{ name: 'Fresh', source: 'x' }]);
		await handlers(d).handler('xlide.restoreModuleFromHead')({ filePath: PROJECT, moduleName: 'Fresh' });
		expect(vi.mocked(vscode.window.showInformationMessage).mock.calls[0][0]).toContain('added since the last commit');
		expect(ui.openTextDocument).not.toHaveBeenCalled();
	});
});

describe('gitChangesReport', () => {
	it('reports each changed module with a unified diff against the revision', async () => {
		const d = deps(
			[{ name: 'Same', source: 'x' }, { name: 'Changed', source: 'Sub A()\r\n    x = 1\r\nEnd Sub\r\n' }, { name: 'Gone', source: 'g' }],
			[{ name: 'Same', source: 'x' }, { name: 'Changed', source: 'Sub A()\r\n    x = 2\r\nEnd Sub\r\n' }, { name: 'Fresh', source: 'f' }],
		);
		const report = await gitChangesReport(d, PROJECT);
		expect(report.tracked).toBe(true);
		if (!report.tracked) { return; }
		expect(report.head).toBe('HEAD');
		expect(report.changes.map((change) => `${change.name}:${change.kind}`)).toEqual(['Changed:modified', 'Fresh:added', 'Gone:removed']);
		expect(report.changes[0].diff).toBe([
			'--- Changed (HEAD)',
			'+++ Changed (current)',
			'@@ -1,3 +1,3 @@',
			' Sub A()',
			'-    x = 1',
			'+    x = 2',
			' End Sub',
		].join('\n'));
		expect(report.changes[1].diff).toContain('+f');
		expect(report.changes[2].diff).toContain('-g');
	});

	it('narrows to one module when asked', async () => {
		const d = deps([{ name: 'A', source: '1' }, { name: 'B', source: '1' }], [{ name: 'A', source: '2' }, { name: 'B', source: '2' }]);
		const report = await gitChangesReport(d, PROJECT, 'HEAD', 'b');
		expect(report.changes.map((change) => change.name)).toEqual(['B']);
	});

	it('answers with a reason instead of throwing when git has no say', async () => {
		const outside = await gitChangesReport(deps([], [], fakeGit({ inRepo: false })), PROJECT);
		expect(outside).toMatchObject({ tracked: false, reason: expect.stringContaining('not inside a git repository'), changes: [] });
		const untracked = await gitChangesReport(deps([], [], fakeGit({ tracked: false })), PROJECT);
		expect(untracked).toMatchObject({ tracked: false, reason: expect.stringContaining('not tracked') });
		const absent = await gitChangesReport(deps(undefined, []), PROJECT, 'abc1234');
		expect(absent).toMatchObject({ tracked: false, reason: expect.stringContaining('not in abc1234') });
	});
});

describe('Show Module History', () => {
	const second = { hash: 'aaaa1111', shortHash: 'aaaa111', author: 'Ann', date: '2026-09-01', subject: 'second' };
	const first = { hash: 'bbbb2222', shortHash: 'bbbb222', author: 'Bob', date: '2026-08-01', subject: 'first' };
	const ref = { root: ROOT, relativePath: 'Book.xlsm' };

	/** A history in which the module text depends on the commit. */
	function historyDeps(byHash: Record<string, ModuleSnapshot[] | undefined>) {
		const d = deps([], []);
		d.modulesAtRevision = vi.fn(async (_project: string, _ref: unknown, revision: string) => byHash[revision]);
		return d;
	}

	it('keeps the commits where the module text moved, newest first, each against the one before', async () => {
		const d = historyDeps({
			aaaa1111: [{ name: 'Module1', source: 'v2' }],
			bbbb2222: [{ name: 'Module1', source: 'v1' }],
		});
		const entries = await moduleHistoryEntries(d, PROJECT, ref, 'module1', [second, first], { windowComplete: true });
		expect(entries.map((entry) => [entry.commit.shortHash, entry.kind, entry.previous, entry.source])).toEqual([
			['aaaa111', 'modified', 'v1', 'v2'],
			['bbbb222', 'added', undefined, 'v1'],
		]);
		expect(entries[0].previousCommit).toBe(first);
	});

	it('says nothing about the oldest commit when the window was cut short', async () => {
		const d = historyDeps({
			aaaa1111: [{ name: 'Module1', source: 'same' }],
			bbbb2222: [{ name: 'Module1', source: 'same' }],
		});
		expect(await moduleHistoryEntries(d, PROJECT, ref, 'Module1', [second, first], { windowComplete: false })).toEqual([]);
	});

	it('reports a module a commit removed, and one it added', async () => {
		const removed = historyDeps({ aaaa1111: [], bbbb2222: [{ name: 'Module1', source: 'v1' }] });
		expect((await moduleHistoryEntries(removed, PROJECT, ref, 'Module1', [second, first], { windowComplete: false }))
			.map((entry) => entry.kind)).toEqual(['removed']);
		// A commit's content never changes, so its snapshot is cached by hash;
		// this test reuses the hashes with other content.
		resetGitSnapshotCacheForTests();
		const added = historyDeps({ aaaa1111: [{ name: 'Module1', source: 'v1' }], bbbb2222: [] });
		expect((await moduleHistoryEntries(added, PROJECT, ref, 'Module1', [second, first], { windowComplete: false }))
			.map((entry) => entry.kind)).toEqual(['added']);
	});

	it('reads each commit once and answers the next question from memory', async () => {
		const d = historyDeps({
			aaaa1111: [{ name: 'Module1', source: 'v2' }],
			bbbb2222: [{ name: 'Module1', source: 'v1' }],
		});
		await moduleHistoryEntries(d, PROJECT, ref, 'Module1', [second, first], { windowComplete: true });
		await moduleHistoryEntries(d, PROJECT, ref, 'Module1', [second, first], { windowComplete: true });
		expect(d.modulesAtRevision).toHaveBeenCalledTimes(2);
	});

	it('stops early and answers nothing when cancelled', async () => {
		const d = historyDeps({ aaaa1111: [{ name: 'Module1', source: 'v2' }], bbbb2222: [{ name: 'Module1', source: 'v1' }] });
		expect(await moduleHistoryEntries(d, PROJECT, ref, 'Module1', [second, first], { windowComplete: true, isCancelled: () => true })).toEqual([]);
	});

	it('offers the changing commits and opens the diff of the pick, before against in', async () => {
		const d = historyDeps({
			aaaa1111: [{ name: 'Module1', source: 'v2' }],
			bbbb2222: [{ name: 'Module1', source: 'v1' }],
		});
		ui.quickPick.mockImplementationOnce(async (items: unknown[]) => (items as Array<{ label: string }>).find((item) => item.label.includes('second')));
		const { handler, provider } = handlers(d);
		await handler('xlide.moduleHistory')({ filePath: PROJECT, moduleName: 'Module1' });

		const offered = ui.quickPick.mock.calls[0][0] as Array<{ label: string; detail: string }>;
		expect(offered.map((item) => `${item.label} ${item.detail}`)).toEqual([
			'$(diff-modified) aaaa111  second modified',
			'$(diff-added) bbbb222  first added',
		]);
		const [diff] = diffCalls();
		expect(diff.title).toBe('Module1: bbbb222 ↔ aaaa111');
		expect(provider.provideTextDocumentContent(diff.left)).toBe('v1');
		expect(provider.provideTextDocumentContent(diff.right)).toBe('v2');
	});

	it('opens a named commit without the picker', async () => {
		const d = historyDeps({
			aaaa1111: [{ name: 'Module1', source: 'v2' }],
			bbbb2222: [{ name: 'Module1', source: 'v1' }],
		});
		const { handler } = handlers(d);
		await handler('xlide.moduleHistory')({ filePath: PROJECT, moduleName: 'Module1' }, { commit: 'bbbb' });
		expect(ui.quickPick).not.toHaveBeenCalled();
		expect(diffCalls()[0].title).toBe('Module1: added in bbbb222');
	});

	it('says so when the module did not change in the window', async () => {
		const d = historyDeps({
			aaaa1111: [{ name: 'Module1', source: 'same' }],
			bbbb2222: [{ name: 'Module1', source: 'same' }],
		});
		await handlers(d).handler('xlide.moduleHistory')({ filePath: PROJECT, moduleName: 'Module1' }, { limit: 2 });
		expect(vi.mocked(vscode.window.showInformationMessage).mock.calls[0][0]).toContain('did not change in the last 2 commits');
		expect(diffCalls()).toHaveLength(0);
	});
});

describe('moduleChanges', () => {
	it('compares by name without case and by text without line endings', () => {
		const committed = new Map<string, ModuleSnapshot>([['a', { name: 'A', source: 'x\n' }], ['b', { name: 'B', source: '1' }]]);
		const current = new Map<string, ModuleSnapshot>([['a', { name: 'a', source: 'x\r\n' }], ['c', { name: 'C', source: '' }]]);
		expect(moduleChanges(committed, current)).toEqual([
			{ name: 'B', kind: 'removed' },
			{ name: 'C', kind: 'added' },
		]);
	});
});
