// Format All Modules: counts first, asks, formats each module on its own
// document, saves the ones that were clean, and reports.

import { beforeEach, describe, expect, it, vi } from 'vitest';

const ui = vi.hoisted(() => ({
	openTextDocument: vi.fn(),
	applyEdit: vi.fn(async () => true),
	showInformationMessage: vi.fn(async () => 'Format'),
	withProgress: vi.fn((_options: unknown, task: (progress: { report(): void }, token: { isCancellationRequested: boolean }) => unknown) =>
		task({ report: () => undefined }, { isCancellationRequested: false })),
	setStatusBarMessage: vi.fn(),
	executeCommand: vi.fn(),
	registerCommand: vi.fn(() => ({ dispose: () => undefined })),
}));

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
	WorkspaceEdit: class {
		entries: Array<[unknown, unknown]> = [];
		set(uri: unknown, edits: unknown): void { this.entries.push([uri, edits]); }
	},
	ProgressLocation: { Notification: 15 },
	commands: { executeCommand: ui.executeCommand, registerCommand: ui.registerCommand },
	window: {
		showInformationMessage: ui.showInformationMessage,
		withProgress: ui.withProgress,
		setStatusBarMessage: ui.setStatusBarMessage,
		activeTextEditor: undefined,
	},
	workspace: {
		openTextDocument: ui.openTextDocument,
		applyEdit: ui.applyEdit,
		getConfiguration: vi.fn(() => ({ get: (_key: string, fallback?: unknown) => fallback, inspect: () => ({}) })),
	},
}));

import { registerFormatCommands } from '../src/commands/formatCommands';
import type { CommandDeps } from '../src/commands/shared';

const PROJECT = process.platform === 'win32' ? 'C:\\work\\Book.xlsm' : '/work/Book.xlsm';

interface FakeDocument {
	uri: { path: string; toString(): string };
	version: number;
	isDirty: boolean;
	save: ReturnType<typeof vi.fn>;
}

function fakeDocument(name: string, dirty = false): FakeDocument {
	return {
		uri: { path: `/work/Book.xlsm/${name}.bas`, toString: () => `xlide-vba:${name}` },
		version: 1,
		isDirty: dirty,
		save: vi.fn(async () => true),
	};
}

function setUp(modules: Array<{ name: string; source?: string }>, edits: Record<string, unknown[]>, documents: Record<string, FakeDocument>) {
	const out = { appendLine: vi.fn() };
	const deps = {
		bridge: { call: vi.fn(async (method: string) => (method === 'readModules' ? modules : [])) },
		out,
	} as unknown as CommandDeps;
	ui.openTextDocument.mockImplementation(async (uri: { path: string }) => {
		const name = uri.path.split('/').pop()!.replace(/\.bas$/, '');
		return documents[name] ?? fakeDocument(name);
	});
	ui.executeCommand.mockImplementation(async (command: string, uri: { path: string }) => {
		if (command !== 'vscode.executeFormatDocumentProvider') {
			return undefined;
		}
		const name = uri.path.split('/').pop()!.replace(/\.bas$/, '');
		return edits[name] ?? [];
	});
	registerFormatCommands(deps);
	const call = ui.registerCommand.mock.calls.find(([id]) => id === 'xlide.formatAllModules')!;
	const handler = call[1] as (node?: unknown, options?: unknown) => Promise<void>;
	return { handler, out };
}

beforeEach(() => {
	ui.openTextDocument.mockReset();
	ui.applyEdit.mockClear();
	ui.showInformationMessage.mockClear();
	ui.withProgress.mockClear();
	ui.setStatusBarMessage.mockClear();
	ui.executeCommand.mockReset();
	ui.registerCommand.mockClear();
});

describe('Format All Modules', () => {
	it('formats the modules with edits, saves the clean ones, and leaves a dirty one open', async () => {
		const dirty = fakeDocument('Dirty', true);
		const clean = fakeDocument('Clean');
		const { handler } = setUp(
			[{ name: 'Clean', source: 'x' }, { name: 'Dirty', source: 'y' }, { name: 'Tidy', source: 'z' }, { name: 'Design' }],
			{ Clean: [{ newText: 'a' }], Dirty: [{ newText: 'b' }] },
			{ Dirty: dirty, Clean: clean },
		);

		await handler({ filePath: PROJECT }, {});

		expect(ui.showInformationMessage.mock.calls[0][0]).toBe('Format 2 of 3 modules in Book.xlsm?');
		expect(ui.applyEdit).toHaveBeenCalledTimes(2);
		expect(clean.save).toHaveBeenCalledTimes(1);
		expect(dirty.save).not.toHaveBeenCalled();
		expect(ui.setStatusBarMessage.mock.calls[0][0]).toBe('XLIDE: Formatted 2 of 3 modules in Book.xlsm, 1 left unsaved');
		// The design without a module is never opened.
		expect(ui.openTextDocument).toHaveBeenCalledTimes(3);
	});

	it('does nothing but say so when every module is already formatted', async () => {
		const { handler } = setUp([{ name: 'A', source: 'x' }], {}, {});
		await handler({ filePath: PROJECT });
		expect(ui.showInformationMessage).not.toHaveBeenCalled();
		expect(ui.applyEdit).not.toHaveBeenCalled();
		expect(ui.setStatusBarMessage.mock.calls[0][0]).toBe('XLIDE: All 1 modules in Book.xlsm are already formatted');
	});

	it('stops when the confirmation is declined', async () => {
		ui.showInformationMessage.mockResolvedValueOnce(undefined as never);
		const { handler } = setUp([{ name: 'A', source: 'x' }], { A: [{ newText: 'a' }] }, {});
		await handler({ filePath: PROJECT });
		expect(ui.applyEdit).not.toHaveBeenCalled();
	});

	it('skips the confirmation when asked to', async () => {
		const { handler } = setUp([{ name: 'A', source: 'x' }], { A: [{ newText: 'a' }] }, {});
		await handler({ filePath: PROJECT }, { withoutPrompt: true });
		expect(ui.showInformationMessage).not.toHaveBeenCalled();
		expect(ui.applyEdit).toHaveBeenCalledTimes(1);
	});

	it('reports a module the editor refused and formats the rest', async () => {
		ui.applyEdit.mockResolvedValueOnce(false);
		const { handler, out } = setUp(
			[{ name: 'A', source: 'x' }, { name: 'B', source: 'y' }],
			{ A: [{ newText: 'a' }], B: [{ newText: 'b' }] },
			{},
		);
		await handler({ filePath: PROJECT }, { withoutPrompt: true });
		expect(ui.setStatusBarMessage.mock.calls[0][0]).toBe('XLIDE: Formatted 1 of 2 modules in Book.xlsm, 1 failed, see the XLIDE output');
		expect(out.appendLine.mock.calls[0][0]).toContain('[format] A: the editor refused the edit');
	});

	it('asks for a file when neither a row nor a module editor gives one', async () => {
		const { handler } = setUp([], {}, {});
		await handler(undefined);
		expect(ui.showInformationMessage.mock.calls[0][0]).toContain('Select a file');
		expect(ui.openTextDocument).not.toHaveBeenCalled();
	});
});
