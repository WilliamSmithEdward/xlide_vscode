import { vi } from 'vitest';

/**
 * Shared `vscode` module stub for vitest. Each test file still declares its own
 * `vi.mock('vscode', ...)` (vitest hoisting prevents a single global mock), but
 * the factory delegates here so every file shares one stub vocabulary:
 *
 *   vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());
 *
 * Per-file overrides merge into the base shape (one level deep for the
 * `commands`/`env`/`lm`/`window`/`workspace` namespaces):
 *
 *   vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
 *       workspace: { findFiles: myHoistedSpy },
 *   }));
 *
 * Per-test overrides go through the vi.fn() leaves, e.g.
 * `vi.mocked(vscode.window.createWebviewPanel).mockReturnValue(panel)`.
 */

export class Disposable {
	static from(...disposables: Array<{ dispose?: () => void }>): Disposable {
		return new Disposable(() => disposables.forEach((disposable) => disposable.dispose?.()));
	}
	constructor(private readonly callOnDispose: () => void = () => undefined) {}
	dispose(): void { this.callOnDispose(); }
}

/** Functional EventEmitter: `fire` notifies subscribed listeners and records values. */
export class EventEmitter<T> {
	readonly events: T[] = [];
	private readonly listeners = new Set<(value: T) => void>();
	readonly event = (listener: (value: T) => void): Disposable => {
		this.listeners.add(listener);
		return new Disposable(() => this.listeners.delete(listener));
	};
	fire(value: T): void {
		this.events.push(value);
		for (const listener of [...this.listeners]) {
			listener(value);
		}
	}
	dispose(): void { this.listeners.clear(); }
}

export class CancellationError extends Error {}

export class Position {
	constructor(
		public line: number,
		public character: number,
	) {}
}

export class Range {
	constructor(
		public start: Position,
		public end: Position,
	) {}
}

export class Location {
	constructor(
		public uri: unknown,
		public range: Range,
	) {}
}

export class MarkdownString {
	supportThemeIcons = false;
	constructor(public value = '') {}
	appendMarkdown = vi.fn((text: string) => { this.value += text; });
}

export class ThemeIcon {
	constructor(readonly id: string) {}
}

export class TreeItem {
	id?: string;
	iconPath?: unknown;
	tooltip?: unknown;
	description?: string;
	contextValue?: string;
	command?: unknown;
	constructor(
		readonly label: string,
		readonly collapsibleState?: number,
	) {}
}

export class RelativePattern {
	constructor(
		public baseUri: unknown,
		public pattern: string,
	) {}
}

export class TabInputText {
	constructor(readonly uri: unknown) {}
}

export class LanguageModelTextPart {
	constructor(readonly value: string) {}
}

export class LanguageModelToolResult {
	constructor(readonly parts: unknown[]) {}
}

/** Builds the mocked `vscode` module shape, merging `overrides` over the base stubs. */
export function vscodeMock(overrides: Record<string, unknown> = {}): Record<string, unknown> {
	const base: Record<string, Record<string, unknown>> = {
		commands: {
			executeCommand: vi.fn(),
			registerCommand: vi.fn(() => new Disposable()),
		},
		env: {
			clipboard: { writeText: vi.fn() },
		},
		lm: {
			registerTool: vi.fn(() => new Disposable()),
		},
		window: {
			createWebviewPanel: vi.fn(),
			showErrorMessage: vi.fn(async () => undefined),
			showInformationMessage: vi.fn(async () => undefined),
			showSaveDialog: vi.fn(async () => undefined),
			showWarningMessage: vi.fn(async () => undefined),
			tabGroups: { all: [], close: vi.fn(async () => true) },
		},
		workspace: {
			textDocuments: [],
			findFiles: vi.fn(async () => []),
			getConfiguration: vi.fn(() => ({
				get: (_key: string, fallback?: unknown) => fallback,
				inspect: () => ({}),
			})),
			onDidChangeConfiguration: vi.fn(() => new Disposable()),
			onDidChangeTextDocument: vi.fn(() => new Disposable()),
			onDidSaveTextDocument: vi.fn(() => new Disposable()),
			onDidCloseTextDocument: vi.fn(() => new Disposable()),
			onDidOpenTextDocument: vi.fn(() => new Disposable()),
			registerTextDocumentContentProvider: vi.fn(() => new Disposable()),
			createFileSystemWatcher: vi.fn(() => ({
				dispose: vi.fn(),
				onDidCreate: vi.fn(() => new Disposable()),
				onDidChange: vi.fn(() => new Disposable()),
				onDidDelete: vi.fn(() => new Disposable()),
			})),
			fs: { writeFile: vi.fn() },
		},
	};
	const merged: Record<string, unknown> = {
		CancellationError,
		Disposable,
		EventEmitter,
		LanguageModelTextPart,
		LanguageModelToolResult,
		Location,
		MarkdownString,
		Position,
		Range,
		RelativePattern,
		TabInputText,
		ThemeIcon,
		TreeItem,
		FileChangeType: { Changed: 1, Created: 2, Deleted: 3 },
		FileType: { Unknown: 0, File: 1, Directory: 2 },
		TreeItemCollapsibleState: { None: 0, Collapsed: 1, Expanded: 2 },
		ViewColumn: { Active: -1, Beside: -2, One: 1, Two: 2 },
		FileSystemError: {
			FileNotFound: (message?: string) => new Error(message),
			NoPermissions: (message?: string) => new Error(message),
			Unavailable: (message?: string) => new Error(message),
		},
		Uri: {
			file: (fsPath: string) => ({
				scheme: 'file',
				fsPath,
				path: fsPath,
				toString: () => fsPath,
			}),
			parse: (value: string) => ({
				scheme: value.split(':', 1)[0],
				path: value.includes(':') ? value.slice(value.indexOf(':') + 1) : value,
				toString: () => value,
			}),
			from: (components: { scheme?: string; authority?: string; path?: string }) => ({
				scheme: components.scheme ?? '',
				authority: components.authority ?? '',
				path: components.path ?? '',
				fsPath: components.path ?? '',
				toString: () => `${components.scheme ?? ''}:${components.path ?? ''}`,
			}),
		},
		...base,
		...overrides,
	};
	for (const namespace of ['commands', 'env', 'lm', 'window', 'workspace']) {
		const override = overrides[namespace];
		if (override && typeof override === 'object') {
			merged[namespace] = { ...base[namespace], ...override };
		}
	}
	return merged;
}
