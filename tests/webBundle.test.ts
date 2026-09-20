import { describe, expect, it } from 'vitest';
import * as esbuild from 'esbuild';
import * as fs from 'fs';
import * as path from 'path';
import { readModulesFromBuffer } from '../src/vba/projectService';
// eslint-disable-next-line @typescript-eslint/no-var-requires
const { webLeafSwap, webInject, FORBIDDEN_BUILTINS, ROOT } = require('../webBuild.js');

// A tripwire, not a build step. XLIDE's container engine is dependency-free
// TypeScript that reaches Node only through the leaf modules webBuild.js
// swaps, which is what makes a browser build possible at all. Nothing in the
// engine stops someone adding `import * as fs` to a reader six months from
// now, and the desktop build would not notice. This does.

/** Bundles one entry the way the browser build will, and reports the result. */
async function bundleForBrowser(entry: string, minify = false): Promise<esbuild.BuildResult> {
	return esbuild.build({
		entryPoints: [path.join(ROOT, entry)],
		bundle: true,
		format: 'cjs',
		platform: 'browser',
		write: false,
		logLevel: 'silent',
		external: ['vscode'],
		minify,
		inject: webInject(),
		plugins: [webLeafSwap()],
	});
}

/** Bundles a virtual entry, for reaching several exports at once. */
async function bundleSource(contents: string): Promise<esbuild.BuildResult> {
	return esbuild.build({
		stdin: { contents, resolveDir: ROOT, loader: 'ts' },
		bundle: true,
		format: 'cjs',
		platform: 'browser',
		write: false,
		logLevel: 'silent',
		external: ['vscode'],
		inject: webInject(),
		plugins: [webLeafSwap()],
	});
}

/**
 * A `vscode` whose workspace.fs is an in-memory filesystem, for driving the
 * browser build's engine path the way github.dev would.
 */
function fakeVscodeWithFiles(files: Map<string, Buffer>): Record<string, unknown> {
	const stamps = new Map<string, number>();
	for (const key of files.keys()) {
		stamps.set(key, 1000);
	}

	class FileSystemError extends Error {
		constructor(readonly code: string, message: string) {
			super(message);
		}
		static FileNotFound(p: string): FileSystemError {
			return new FileSystemError('FileNotFound', p);
		}
	}

	const reads: string[] = [];
	const fs = {
		reads,
		async stat(uri: { fsPath: string }) {
			const data = files.get(uri.fsPath);
			if (!data) {
				throw FileSystemError.FileNotFound(uri.fsPath);
			}
			return { type: 1, ctime: 0, mtime: stamps.get(uri.fsPath) ?? 0, size: data.length };
		},
		async readFile(uri: { fsPath: string }) {
			const data = files.get(uri.fsPath);
			if (!data) {
				throw FileSystemError.FileNotFound(uri.fsPath);
			}
			reads.push(uri.fsPath);
			return new Uint8Array(data);
		},
		async writeFile(uri: { fsPath: string }, bytes: Uint8Array) {
			files.set(uri.fsPath, Buffer.from(bytes));
			stamps.set(uri.fsPath, (stamps.get(uri.fsPath) ?? 0) + 1000);
		},
	};

	return {
		...fakeVscode(),
		FileSystemError,
		FileType: { Unknown: 0, File: 1, Directory: 2, SymbolicLink: 64 },
		Uri: { file: (p: string) => ({ fsPath: p, path: p, scheme: 'file', toString: () => p }) },
		workspace: {
			fs,
			// The workspace folder the URI resolver rebuilds paths against,
			// shaped like test-web's mount: a root whose path is '/'.
			workspaceFolders: [
				{ uri: { fsPath: '\\', path: '/', scheme: 'file', with: withOverride('/'), toString: () => '/' } },
			],
			getConfiguration: () => ({ get: (_k: string, fallback?: unknown) => fallback }),
			createFileSystemWatcher: () => ({
				onDidChange: () => ({ dispose: () => undefined }),
				onDidCreate: () => ({ dispose: () => undefined }),
				onDidDelete: () => ({ dispose: () => undefined }),
				dispose: () => undefined,
			}),
		},
		RelativePattern: class { constructor(public base: unknown, public pattern: string) {} },
	};
}

/**
 * vscode.Uri.with(), enough for workspaceUriFor to rebuild a path. fsPath is
 * left identical to path so it keys the in-memory filesystem the same way
 * Uri.file() does above.
 */
function withOverride(basePath: string) {
	return function (this: unknown, change: { path?: string }) {
		const next = change.path ?? basePath;
		return { fsPath: next, path: next, scheme: 'file', toString: () => next };
	};
}

/** Just enough of the `vscode` module for a bundle to load and run. */
function fakeVscode(): Record<string, unknown> {
	class EventEmitter {
		fire(): void { /* nothing listens in a test */ }
		get event() {
			return () => ({ dispose: () => undefined });
		}
		dispose(): void { /* nothing to release */ }
	}
	class Disposable {
		constructor(private readonly _onDispose: () => void) {}
		dispose(): void {
			this._onDispose();
		}
	}
	return { EventEmitter, Disposable, workspace: {}, window: {}, commands: {} };
}

/**
 * Runs a CommonJS bundle with the Node globals shadowed away, so anything
 * still reaching for one throws here rather than in someone's browser.
 */
function runWithoutNodeGlobals(
	code: string,
	vscodeModule: Record<string, unknown> = fakeVscode(),
): Record<string, unknown> {
	const module = { exports: {} as Record<string, unknown> };
	const require = (name: string): unknown => {
		if (name === 'vscode') {
			return vscodeModule;
		}
		throw new Error(`the web bundle required ${name}`);
	};
	// eslint-disable-next-line no-new-func
	const run = new Function(
		'module',
		'exports',
		'require',
		'Buffer',
		'process',
		'__dirname',
		'__filename',
		code,
	);
	run(module, module.exports, require, undefined, undefined, undefined, undefined);
	return module.exports;
}

function bundleText(result: esbuild.BuildResult): string {
	const file = result.outputFiles?.[0];
	expect(file, 'esbuild produced no output').toBeDefined();
	return Buffer.from(file!.contents).toString('utf8');
}

describe('the container engine bundles for a browser', () => {
	it('bundles every macro container format', async () => {
		const result = await bundleForBrowser('src/vba/macroContainer.ts');
		expect(result.errors).toEqual([]);
		expect(bundleText(result).length).toBeGreaterThan(1000);
	}, 60000);

	it('bundles the VBA project reader and writer', async () => {
		const result = await bundleForBrowser('src/vba/vbaProject.ts');
		expect(result.errors).toEqual([]);
	}, 60000);

	it('bundles the project service, the engine entry the extension calls', async () => {
		const result = await bundleForBrowser('src/vba/projectService.ts');
		expect(result.errors).toEqual([]);
	}, 60000);

	it('bundles the VB6 project reader', async () => {
		const result = await bundleForBrowser('src/vba/vb6/vb6Project.ts');
		expect(result.errors).toEqual([]);
	}, 60000);

	it('bundles the analyzer', async () => {
		const result = await bundleForBrowser('src/analyzer/index.ts');
		expect(result.errors).toEqual([]);
	}, 60000);

	it.each(['src/vba/macroContainer.ts', 'src/vba/projectService.ts', 'src/analyzer/index.ts'])(
		'pulls in no Node builtin: %s',
		async (entry) => {
			const text = bundleText(await bundleForBrowser(entry));
			for (const builtin of FORBIDDEN_BUILTINS) {
				expect(text.includes(`require("${builtin}")`), `${entry} requires ${builtin}`).toBe(false);
				expect(
					text.includes(`require("node:${builtin}")`),
					`${entry} requires node:${builtin}`,
				).toBe(false);
			}
		},
		60000,
	);

	it('carries the pure-TypeScript codec, not zlib', async () => {
		const text = bundleText(await bundleForBrowser('src/vba/macroContainer.ts'));
		expect(text).toContain('pure-ts');
		expect(text).not.toContain('inflateRawSync');
	}, 60000);
});

describe('the browser bundle actually runs', () => {
	const FIXTURE = path.join(__dirname, 'fixtures', 'binaries', 'ShapesFixture.xlsm');

	it('reads a real workbook with no Node globals available', async () => {
		const result = await bundleSource(
			[
				"export { readModulesFromBuffer } from './src/vba/projectService';",
				"export { WebBuffer as Buffer } from './src/util/webBuffer';",
			].join('\n'),
		);
		expect(result.errors).toEqual([]);

		const api = runWithoutNodeGlobals(bundleText(result)) as {
			readModulesFromBuffer: (data: unknown, full: boolean) => { name: string; source: string }[];
			Buffer: { from(bytes: Uint8Array): unknown };
		};

		// Hand it the bytes through its OWN Buffer, the way the extension
		// layer will after reading them from the workspace.
		const bytes = new Uint8Array(fs.readFileSync(FIXTURE));
		const inBrowser = api.readModulesFromBuffer(api.Buffer.from(bytes), true);

		const onDesktop = readModulesFromBuffer(fs.readFileSync(FIXTURE), true);

		expect(inBrowser.map((m) => m.name)).toEqual(onDesktop.map((m) => m.name));
		for (let i = 0; i < onDesktop.length; i++) {
			expect(inBrowser[i].source, `${onDesktop[i].name} differs`).toBe(onDesktop[i].source);
		}
		expect(onDesktop.length).toBeGreaterThan(0);
	}, 60000);
});

describe('the whole extension bundles for a browser', () => {
	// The build package.json points `browser` at. If this fails, the web
	// build is broken however green everything else is.
	it('bundles src/extension.ts with no Node builtin, no process and no __dirname', async () => {
		const result = await bundleForBrowser('src/extension.ts');
		expect(result.errors).toEqual([]);

		const text = bundleText(result);
		for (const builtin of FORBIDDEN_BUILTINS) {
			expect(text.includes(`require("${builtin}")`), `requires ${builtin}`).toBe(false);
			expect(text.includes(`require("node:${builtin}")`), `requires node:${builtin}`).toBe(false);
		}
		// `process` and `__dirname` are globals rather than imports, so esbuild
		// cannot report them: they would only fail once someone opened the
		// extension in a browser.
		expect(text.match(/\bprocess\.[a-zA-Z]+/g) ?? [], 'process is not available in a browser')
			.toEqual([]);
		expect(text.match(/\b__dirname\b/g) ?? [], '__dirname is not available in a browser')
			.toEqual([]);
	}, 120000);

	it('leaves the desktop-only modules out of it', async () => {
		// Minified, so a doc comment mentioning powershell.exe cannot pass for
		// code that spawns it.
		const text = bundleText(await bundleForBrowser('src/extension.ts', true));
		// String literals that exist only inside modules platformFeaturesNode
		// reaches. Any of them here means a desktop-only module came along.
		for (const marker of [
			'powershell.exe',          // the PowerShell spawner
			'run-vba-tests.ps1',       // the VBA test host script
			'taskkill.exe',            // the process-tree kill
			'XlideTestModalWatcher',   // the test host's modal watcher
		]) {
			expect(text.includes(marker), `web bundle carries ${marker}`).toBe(false);
		}
	}, 120000);
});

describe('the browser reads and writes a workbook through workspace.fs', () => {
	// The whole point of the port: load the container through the editor's
	// filesystem, run the synchronous engine over it, and write back what it
	// produced. Nothing here touches a real disk except to seed the fixture.
	const FIXTURE = path.join(__dirname, 'fixtures', 'binaries', 'ShapesFixture.xlsm');
	const WORKSPACE_PATH = '/repo/Book.xlsm';

	async function loadEngine(files: Map<string, Buffer>) {
		const result = await bundleSource(
			[
				"export { enginePriming } from './src/enginePrimingWeb';",
				"export * as svc from './src/vba/projectService';",
			].join('\n'),
		);
		expect(result.errors).toEqual([]);
		const vscodeModule = fakeVscodeWithFiles(files);
		const loaded = runWithoutNodeGlobals(bundleText(result), vscodeModule) as Record<
			string,
			unknown
		>;
		loaded.reads = (vscodeModule.workspace as { fs: { reads: string[] } }).fs.reads;
		return loaded as {
			reads: string[];
			enginePriming: {
				prime(paths: string[]): Promise<void>;
				flush(): Promise<void>;
				discard(): void;
			};
			svc: {
				listModules(p: string): { name: string }[];
				readModule(p: string, m: string): { source: string };
				writeModule(p: string, m: string, source: string): { ok: boolean };
			};
		};
	}

	it('lists and reads modules after priming', async () => {
		const files = new Map([[WORKSPACE_PATH, fs.readFileSync(FIXTURE)]]);
		const { enginePriming, svc } = await loadEngine(files);

		await enginePriming.prime([WORKSPACE_PATH]);
		const names = svc.listModules(WORKSPACE_PATH).map((m) => m.name).sort();

		const onDesktop = readModulesFromBuffer(fs.readFileSync(FIXTURE), false)
			.map((m) => m.name)
			.sort();
		expect(names).toEqual(onDesktop);
		expect(names.length).toBeGreaterThan(0);
	}, 60000);

	it('writes a module back into the workspace, and Node reads it', async () => {
		const files = new Map([[WORKSPACE_PATH, fs.readFileSync(FIXTURE)]]);
		const original = files.get(WORKSPACE_PATH)!;
		const { enginePriming, svc } = await loadEngine(files);

		await enginePriming.prime([WORKSPACE_PATH]);
		const target = svc.listModules(WORKSPACE_PATH)[0].name;
		svc.writeModule(
			WORKSPACE_PATH,
			target,
			'Sub WrittenInTheBrowser()\r\n    Debug.Print "hello"\r\nEnd Sub\r\n',
		);
		await enginePriming.flush();

		// The workspace now holds different bytes...
		const saved = files.get(WORKSPACE_PATH)!;
		expect(saved.equals(original)).toBe(false);

		// ...and they are a workbook the ordinary desktop engine reads.
		const modules = readModulesFromBuffer(saved, true);
		expect(modules.map((m) => m.name).sort()).toEqual(
			readModulesFromBuffer(original, true).map((m) => m.name).sort(),
		);
		expect(modules.find((m) => m.name === target)?.source).toContain('WrittenInTheBrowser');
	}, 60000);

	it('does not write anything when the call never flushes', async () => {
		const files = new Map([[WORKSPACE_PATH, fs.readFileSync(FIXTURE)]]);
		const original = Buffer.from(files.get(WORKSPACE_PATH)!);
		const { enginePriming, svc } = await loadEngine(files);

		await enginePriming.prime([WORKSPACE_PATH]);
		svc.writeModule(WORKSPACE_PATH, svc.listModules(WORKSPACE_PATH)[0].name, 'Sub A()\r\nEnd Sub\r\n');
		enginePriming.discard();
		await enginePriming.flush();

		expect(files.get(WORKSPACE_PATH)!.equals(original)).toBe(true);
	}, 60000);

	it('re-reads only when the workspace copy actually changed', async () => {
		const files = new Map([[WORKSPACE_PATH, fs.readFileSync(FIXTURE)]]);
		const { enginePriming, svc } = await loadEngine(files);

		await enginePriming.prime([WORKSPACE_PATH]);
		const before = svc.listModules(WORKSPACE_PATH).length;

		// Priming again with nothing changed must not disturb what is loaded.
		await enginePriming.prime([WORKSPACE_PATH]);
		expect(svc.listModules(WORKSPACE_PATH).length).toBe(before);

		// A write bumps the stamp, so the next prime picks the new bytes up.
		svc.writeModule(
			WORKSPACE_PATH,
			svc.listModules(WORKSPACE_PATH)[0].name,
			'Sub Changed()\r\nEnd Sub\r\n',
		);
		await enginePriming.flush();
		await enginePriming.prime([WORKSPACE_PATH]);
		const target = svc.listModules(WORKSPACE_PATH)[0].name;
		expect(svc.readModule(WORKSPACE_PATH, target).source).toContain('Changed');
	}, 60000);

	it('does not re-read the container it just wrote', async () => {
		const files = new Map([[WORKSPACE_PATH, fs.readFileSync(FIXTURE)]]);
		const { enginePriming, svc, reads } = await loadEngine(files);

		await enginePriming.prime([WORKSPACE_PATH]);
		expect(reads.length).toBe(1);

		svc.writeModule(
			WORKSPACE_PATH,
			svc.listModules(WORKSPACE_PATH)[0].name,
			'Sub Written()\r\nEnd Sub\r\n',
		);
		await enginePriming.flush();

		// The bytes in hand are the bytes just written, so the next call
		// should not pull a whole workbook back over the wire.
		await enginePriming.prime([WORKSPACE_PATH]);
		expect(reads.length, 'the container was re-read after its own write').toBe(1);
		expect(svc.readModule(WORKSPACE_PATH, svc.listModules(WORKSPACE_PATH)[0].name).source)
			.toContain('Written');
	}, 60000);

	it('saves a module through the wrapper the editor actually calls', async () => {
		// One layer up from the engine: runWriteWithHostCoordination is what
		// XlideFileSystemProvider.writeFile calls on every save, and it stats
		// the container BEFORE the write - i.e. before priming has happened.
		// On a desktop that reads the real file; in a browser it must degrade
		// rather than throw HostFileNotPrimedError and lose the save.
		const files = new Map([[WORKSPACE_PATH, fs.readFileSync(FIXTURE)]]);
		const original = Buffer.from(files.get(WORKSPACE_PATH)!);

		const result = await bundleSource(
			[
				"export { enginePriming } from './src/enginePrimingWeb';",
				"export { runWriteWithHostCoordination } from './src/officeWriteCoordinator';",
				"export * as svc from './src/vba/projectService';",
			].join('\n'),
		);
		expect(result.errors).toEqual([]);

		const api = runWithoutNodeGlobals(bundleText(result), fakeVscodeWithFiles(files)) as {
			enginePriming: { prime(p: string[]): Promise<void>; flush(): Promise<void> };
			runWriteWithHostCoordination<T>(p: string, write: () => Promise<T>): Promise<T>;
			svc: {
				listModules(p: string): { name: string }[];
				writeModule(p: string, m: string, source: string): { ok: boolean };
			};
		};

		await api.enginePriming.prime([WORKSPACE_PATH]);
		const target = api.svc.listModules(WORKSPACE_PATH)[0].name;

		// The shape of a real save: the coordinator wraps the engine call,
		// and the flush is what puts the bytes back in the workspace.
		await api.runWriteWithHostCoordination(WORKSPACE_PATH, async () => {
			api.svc.writeModule(
				WORKSPACE_PATH,
				target,
				'Sub SavedThroughTheCoordinator()\r\nEnd Sub\r\n',
			);
			await api.enginePriming.flush();
			return undefined;
		});

		const saved = files.get(WORKSPACE_PATH)!;
		expect(saved.equals(original), 'the workspace copy never changed').toBe(false);

		const modules = readModulesFromBuffer(saved, true);
		expect(modules.map((m) => m.name).sort())
			.toEqual(readModulesFromBuffer(original, true).map((m) => m.name).sort());
		expect(modules.find((m) => m.name === target)?.source)
			.toContain('SavedThroughTheCoordinator');
	}, 60000);

	it('reports a container that is not in the workspace, rather than reading it as empty', async () => {
		const { enginePriming, svc } = await loadEngine(new Map());

		await enginePriming.prime(['/repo/Missing.xlsm']);
		expect(() => svc.listModules('/repo/Missing.xlsm')).toThrow();
	}, 60000);
});

describe('the web platform features', () => {
	// The seam that decides what the browser build does NOT do. Bundling it
	// for a browser is what proves the desktop side - child_process,
	// worker_threads, the test host, the git binary - is genuinely absent
	// rather than merely unused.
	async function loadWebFeatures(): Promise<Record<string, () => unknown>> {
		const result = await bundleForBrowser('src/platformFeaturesWeb.ts');
		expect(result.errors).toEqual([]);
		const exports = runWithoutNodeGlobals(bundleText(result));
		return exports.platformFeatures as Record<string, () => unknown>;
	}

	it('carries the webview templates, which it cannot read from disk', async () => {
		const result = await bundleSource(
			"export { readExtensionTextAsset } from './src/extensionAssetsWeb';",
		);
		expect(result.errors).toEqual([]);
		const api = runWithoutNodeGlobals(bundleText(result)) as {
			readExtensionTextAsset(relativePath: string): string;
		};

		// Every webview template the repo ships, compared with the file the
		// desktop build reads at runtime.
		const dir = path.join(ROOT, 'assets', 'webview');
		const names = fs.readdirSync(dir).filter((name) => fs.statSync(path.join(dir, name)).isFile());
		expect(names.length).toBeGreaterThan(5);
		for (const name of names) {
			const onDisk = fs.readFileSync(path.join(dir, name), 'utf8').replace(/\r\n/g, '\n');
			expect(api.readExtensionTextAsset(`assets/webview/${name}`), name).toBe(onDisk);
		}
	}, 60000);

	it('names the feature, not the mechanism, for an asset it does not carry', async () => {
		const result = await bundleSource(
			"export { readExtensionTextAsset } from './src/extensionAssetsWeb';",
		);
		const api = runWithoutNodeGlobals(bundleText(result)) as {
			readExtensionTextAsset(relativePath: string): string;
		};
		expect(() => api.readExtensionTextAsset('assets/testhost/run-vba-tests.ps1')).toThrow(
			/desktop editor/,
		);
	}, 60000);

	it('bundles with no Node builtin', async () => {
		const text = bundleText(await bundleForBrowser('src/platformFeaturesWeb.ts'));
		for (const builtin of FORBIDDEN_BUILTINS) {
			expect(text.includes(`require("${builtin}")`), `requires ${builtin}`).toBe(false);
			expect(text.includes(`require("node:${builtin}")`), `requires node:${builtin}`).toBe(false);
		}
	}, 60000);

	it('declines every feature that needs a desktop', async () => {
		const features = await loadWebFeatures();
		expect(features.name).toBe('web');
		expect(features.createAnalysisWorker('ignored', () => undefined)).toBeUndefined();
		expect(await features.cleanupStaleTempDirs()).toBeUndefined();
		expect(features.registerAgentTools()).toEqual([]);
		expect(features.registerPlatformCommands({} as never)).toEqual([]);
		expect(features.watchRepositories({} as never)).toEqual([]);
	}, 60000);

	it('gives the tree marks that are absent rather than wrong', async () => {
		const features = await loadWebFeatures();
		const marks = features.createChangeMarks({} as never, () => undefined) as {
			marksFor(p: string): unknown;
			invalidate(p: string): void;
			invalidateAll(): void;
			dispose(): void;
		};
		// undefined is the same answer the desktop gives for a file git has
		// no say over, so the tree already knows to draw no mark.
		expect(marks.marksFor('C:/anything.xlsm')).toBeUndefined();
		expect(() => marks.invalidate('C:/anything.xlsm')).not.toThrow();
		expect(() => marks.invalidateAll()).not.toThrow();
		expect(() => marks.dispose()).not.toThrow();
	}, 60000);
});
