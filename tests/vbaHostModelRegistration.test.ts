import { build } from 'esbuild';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it, vi } from 'vitest';
import { analyzeModule } from '../src/analyzer/diagnostics/analyzeModule';
import type { HostObjectModel } from '../src/analyzer/host/excelObjectModel';

// Only Excel's object model is built into the registry; the extension and its
// worker register the other hosts at load, as tests/registerHostModels.setup.ts
// does for the suite. These check what a caller gets before it registers
// anything, what registering changes, and that a bundle carries a host's model
// only when something registers it.

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');

/** A fresh module graph: the registry as an embedder first sees it. */
async function unregistered() {
	vi.resetModules();
	return {
		registry: await import('../src/analyzer/host/hostRegistry'),
		builtIns: await import('../src/analyzer/host/builtInHostModels'),
		analyzer: await import('../src/analyzer/diagnostics/analyzeModule'),
	};
}

function codes(source: string, diagnostics: ReturnType<typeof analyzeModule>): string[] {
	return diagnostics.map((d) => `${d.code} ${source.slice(d.span.start, d.span.end)}`);
}

function model(typeName: string): () => HostObjectModel {
	return () => ({
		source: 'test',
		hostName: 'Word',
		types: { [typeName]: { displayName: typeName.split('.')[1], members: [] } },
		aliases: {},
		globals: {},
	});
}

/** Which other hosts' model data a bundle of `contents` carries. */
async function hostModelsBundled(contents: string, platform: 'browser' | 'node'): Promise<string[]> {
	const result = await build({
		stdin: { contents, resolveDir: ROOT, loader: 'ts' },
		bundle: true,
		format: 'esm',
		platform,
		external: ['vscode'],
		write: false,
		metafile: true,
		logLevel: 'silent',
	});
	const inputs = Object.values(result.metafile.outputs)[0].inputs;
	return Object.entries(inputs)
		.filter(([file, used]) => /(word|powerpoint|access|vb6)ObjectModelData\.ts$/.test(file) && used.bytesInOutput > 0)
		.map(([file]) => file.replace(/^.*\/(\w+)ObjectModelData\.ts$/, '$1'))
		.sort();
}

const WORD_MODULE = [
	'Option Explicit',
	'Sub Report()',
	'    Dim r As Word.Range',
	'    Set r = ActiveDocument.Content',
	'    r.NoSuchMember',
	'    Documents.Add 1, 2, 3, 4, 5, 6, 7, 8, 9',
	'End Sub',
	'',
].join('\r\n');

const EXCEL_MODULE = [
	'Option Explicit',
	'Sub Report()',
	'    Dim ws As Worksheet',
	'    Set ws = ActiveSheet',
	'    ws.Range("A1").Value = 1 / 0',
	'    ws.NoSuchMember',
	'    Debug.Print Undeclared',
	'End Sub',
	'',
].join('\r\n');

describe('host object model registration', () => {
	it('builds in Excel only: another host answers the empty model until it is registered', async () => {
		const { registry, builtIns } = await unregistered();
		expect(registry.hostObjectModelForToken(undefined)).toBeUndefined();
		expect(registry.hostObjectModelForToken('excel')).toBeUndefined();
		for (const token of ['word', 'powerpoint', 'access', 'vb6']) {
			expect(registry.hostObjectModelForToken(token)).toBe(registry.EMPTY_HOST_MODEL);
		}

		builtIns.registerBuiltInHostModels();
		expect(['word', 'powerpoint', 'access', 'vb6'].map((token) => registry.hostObjectModelForToken(token)?.hostName))
			.toEqual(['Word', 'PowerPoint', 'Access', 'VB6']);
	});

	it('checks a Word module against Word once Word is registered, and says nothing before', async () => {
		const { builtIns, analyzer } = await unregistered();
		expect(codes(WORD_MODULE, analyzer.analyzeModule(WORD_MODULE, { host: 'word' }))).toEqual([]);

		builtIns.registerBuiltInHostModels();
		const registered = codes(WORD_MODULE, analyzer.analyzeModule(WORD_MODULE, { host: 'word' }));
		expect(registered).toEqual(codes(WORD_MODULE, analyzeModule(WORD_MODULE, { host: 'word' })));
		expect(registered).toContain('member-not-found NoSuchMember');
		expect(registered).toContain('argument-count Add');
	});

	it('checks an Excel module the same whether or not the other hosts are registered', async () => {
		const { analyzer } = await unregistered();
		const withoutOthers = codes(EXCEL_MODULE, analyzer.analyzeModule(EXCEL_MODULE, {}));
		expect(withoutOthers.length).toBeGreaterThan(0);
		expect(withoutOthers).toEqual(codes(EXCEL_MODULE, analyzeModule(EXCEL_MODULE, {})));
	});

	it('rebuilds a merged model when a host is registered after it was built', async () => {
		const { registry } = await unregistered();
		registry.registerHostObjectModel('word', model('Word.Before'));
		expect(Object.keys(registry.hostObjectModelForTokens(['excel', 'word'])?.types ?? {})).toContain('Word.Before');

		registry.registerHostObjectModel('word', model('Word.After'));
		const types = Object.keys(registry.hostObjectModelForTokens(['excel', 'word'])?.types ?? {});
		expect(types).toContain('Word.After');
		expect(types).not.toContain('Word.Before');
	});

	it('does not report the host or a referenced library missing when no model for it is registered', async () => {
		const source = 'Option Explicit\r\nPublic Doc As Word.Document\r\n';
		const missing = (diagnostics: ReturnType<typeof analyzeModule>) =>
			diagnostics.filter((d) => d.code === 'missing-library-reference').length;

		// An Excel workbook, as the extension passes one: its host, then its references.
		const { analyzer } = await unregistered();
		expect(missing(analyzer.analyzeModule(source, { host: 'excel' }))).toBe(1);
		expect(missing(analyzer.analyzeModule(source, { host: 'excel', referencedHosts: ['word'] }))).toBe(0);
		// A Word document always has its own library, whatever else it references.
		expect(missing(analyzer.analyzeModule(source, { host: 'word', referencedHosts: ['access'] }))).toBe(0);
		expect(missing(analyzer.analyzeModule(source, { host: 'word', referencedHosts: ['excel'] }))).toBe(0);

		// With every host registered, as in the extension, the answers are the same.
		expect(missing(analyzeModule(source, { host: 'excel' }))).toBe(1);
		expect(missing(analyzeModule(source, { host: 'excel', referencedHosts: ['word'] }))).toBe(0);
		expect(missing(analyzeModule(source, { host: 'word', referencedHosts: ['access'] }))).toBe(0);
		expect(missing(analyzeModule(source, { host: 'word', referencedHosts: ['excel'] }))).toBe(0);
	});

	it('bundles another host\'s model only where something registers it', async () => {
		expect(await hostModelsBundled([
			"export { ProjectIndex, analyzeModule } from './src/analyzer';",
			"export { resolveMemberCompletions } from './src/analyzer/completion/memberAccess';",
		].join('\n'), 'browser')).toEqual([]);

		const all = ['access', 'powerpoint', 'vb6', 'word'];
		expect(await hostModelsBundled("import './src/analysisWorker';", 'node')).toEqual(all);
		expect(await hostModelsBundled("export { activate } from './src/extension';", 'node')).toEqual(all);
	}, 60_000);
});
