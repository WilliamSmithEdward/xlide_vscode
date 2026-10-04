import { afterEach, describe, expect, it, vi } from 'vitest';

const semanticEvents = vi.hoisted(() => ({ close: [] as Array<(document: import('vscode').TextDocument) => void> }));

vi.mock('vscode', async () => ({
	...(await import('./helpers/vscodeMock')).vscodeMock({
        workspace: {
            onDidCloseTextDocument: (listener: (document: import('vscode').TextDocument) => void) => {
                semanticEvents.close.push(listener);
                return { dispose() { const index = semanticEvents.close.indexOf(listener); if (index >= 0) semanticEvents.close.splice(index, 1); } };
            },
        },
        window: {
            onDidChangeTextEditorSelection: vi.fn(() => ({ dispose() {} })),
            onDidChangeActiveTextEditor: vi.fn(() => ({ dispose() {} })),
        },
    }),
	SemanticTokensLegend: class {},
	SemanticTokensBuilder: class {
		rows: unknown[][] = [];
		push(...args: unknown[]) { this.rows.push(args); }
		build() { return { data: new Uint32Array(this.rows.length), rows: this.rows }; }
	},
	DocumentHighlight: class { constructor(public range: unknown, public kind: number) {} },
	DocumentHighlightKind: { Read: 1, Write: 2 },
	SignatureInformation: class { constructor(public label: string) {} },
	SignatureHelp: class {},
	ParameterInformation: class { constructor(public label: string) {} },
}));
vi.mock('../src/vbaDocumentLocation', () => ({
	moduleLocationOfDocument: (document: { uri: { path: string } }) => ({ projectPath: document.uri.path.split('/')[1], moduleName: 'Caller', native: false }),
	moduleDocumentUri: vi.fn(),
}));
vi.mock('../src/vbaDocumentIdentity', () => ({
	analysisSourceForDocument: (document: { getText(): string }) => document.getText(),
	moduleNameFromDocument: () => 'Caller',
	moduleKindFromDocument: () => 'standard',
	liveProjectIndexForDocument: vi.fn(),
	isStandaloneVbaDocument: () => false,
}));
import * as vscode from 'vscode';
import * as analyzer from '../src/analyzer';
import { VbaTypeSemanticTokensProvider } from '../src/vbaSemanticTokensProvider';
import { VbaHoverSignatureProvider } from '../src/vbaHoverSignatureProvider';
import { VbaDocumentHighlightProvider } from '../src/vbaNavigationProviders';
import { buildLiveVbaProjectIndex } from '../src/vbaProjectAnalysis';
import { liveProjectIndexForDocument } from '../src/vbaDocumentIdentity';
import * as references from '../src/vbaReferenceResolution';
import { VbaCaretProcedureTracker } from '../src/vbaCaretProcedure';

function documentFor(source: string, scheme = 'xlide-vba', projectPath = 'Book.xlsm') {
	const lines = source.split('\n');
	const starts = lines.map((_, i) => lines.slice(0, i).reduce((sum, line) => sum + line.length + 1, 0));
	const positionAt = (offset: number) => {
		let line = 0;
		while (line + 1 < starts.length && starts[line + 1] <= offset) line++;
		return new vscode.Position(line, offset - starts[line]);
	};
	const offsetAt = (pos: vscode.Position) => starts[pos.line] + pos.character;
	return {
		version: 1, languageId: 'vba', lineCount: lines.length,
		uri: { scheme, path: '/' + projectPath + '/Caller.bas', toString: () => scheme + ':Caller' },
		getText: vi.fn((range?: vscode.Range) => range ? source.slice(offsetAt(range.start), offsetAt(range.end)) : source),
		lineAt: (line: number) => ({ text: lines[line] }), positionAt, offsetAt,
		getWordRangeAtPosition: (position: vscode.Position) => {
			const start = source.indexOf('Counter', offsetAt(position) - 2);
			return new vscode.Range(positionAt(start), positionAt(start + 'Counter'.length));
		},
	} as unknown as vscode.TextDocument;
}
const active = { isCancellationRequested: false } as vscode.CancellationToken;
const disposables: { dispose(): void }[] = [];
afterEach(() => { for (const item of disposables.splice(0)) item.dispose(); vi.restoreAllMocks(); });

function projectFixture(siblings = 2) {
	const modules = [
		{ moduleName: 'Caller', source: 'Sub Demo()\nCounter = Counter + 1\nEnd Sub\n', type: 'standard' },
		{ moduleName: 'Api', source: 'Public Counter As Long\n', type: 'standard' },
		...Array.from({ length: siblings }, (_, i) => ({
			moduleName: 'Sibling' + i, type: 'standard',
			source: 'Sub Other()\n' + 'Counter = Counter + 1\n'.repeat(40) + 'End Sub\n',
		})),
	];
	const project = buildLiveVbaProjectIndex(modules);
	const byModule = new Map(modules.map(mod => [mod.moduleName.toLowerCase(), mod]));
	const moduleMetadata = new Map(modules.map(mod => [mod.moduleName.toLowerCase(), {
		moduleName: mod.moduleName, moduleKind: 'standard', moduleType: mod.type,
	}]));
	return { modules, project, byModule, moduleMetadata };
}

describe('semantic token provider work', () => {
	it('skips source reads for canceled requests and cache hits', async () => {
		const doc = documentFor('Sub Demo()\nDebug.Print ThisWorkbook.Name\nEnd Sub\n');
		(vscode.workspace as { textDocuments: readonly vscode.TextDocument[] }).textDocuments = [doc];
		const service = { contextForProject: vi.fn() };
		const provider = new VbaTypeSemanticTokensProvider(service as never);
		disposables.push(provider);
		await provider.provideDocumentSemanticTokens(doc, { isCancellationRequested: true } as vscode.CancellationToken);
		expect(doc.getText).not.toHaveBeenCalled();
		const first = await provider.provideDocumentSemanticTokens(doc, active);
		const reads = vi.mocked(doc.getText).mock.calls.length;
		expect(first.data.length).toBeGreaterThan(0);
		expect(await provider.provideDocumentSemanticTokens(doc, active)).toBe(first);
		expect(doc.getText).toHaveBeenCalledTimes(reads);
		expect(service.contextForProject).not.toHaveBeenCalled();
	});

	it('fetches project context once and never asks for diagnostic-only facts', async () => {
		const fixture = projectFixture();
		const service = { contextForProject: vi.fn(async () => fixture) };
		const provider = new VbaTypeSemanticTokensProvider(service as never);
		disposables.push(provider);
		const doc = documentFor(fixture.modules[0].source, 'file');
		(vscode.workspace as { textDocuments: readonly vscode.TextDocument[] }).textDocuments = [doc];
		vi.mocked(liveProjectIndexForDocument).mockResolvedValue(fixture.project);
		const facts = ['writtenNames', 'nameMentions', 'sheetChanges', 'openedFileNumbers'] as const;
		const spies = facts.map(name => vi.spyOn(fixture.project, name));
		await provider.provideDocumentSemanticTokens(doc, active);
		expect(service.contextForProject).toHaveBeenCalledTimes(1);
		for (const spy of spies) expect(spy).not.toHaveBeenCalled();
	});

	it('keeps cached tokens when a stale project-context refresh fails', async () => {
		const fixture = projectFixture();
		const doc = documentFor(fixture.modules[0].source, 'file');
		(vscode.workspace as { textDocuments: readonly vscode.TextDocument[] }).textDocuments = [doc];
		const clock = vi.spyOn(Date, 'now').mockReturnValue(1000);
		const service = { contextForProject: vi.fn(async () => fixture) };
		const provider = new VbaTypeSemanticTokensProvider(service as never);
		disposables.push(provider);
		const first = await provider.provideDocumentSemanticTokens(doc, active);
		clock.mockReturnValue(7000);
		service.contextForProject.mockRejectedValueOnce(new Error('project unavailable'));
		expect(await provider.provideDocumentSemanticTokens(doc, active)).toBe(first);
	});

	it.each([
		['Book.xlsm', { moduleKind: 'userform', moduleType: 'userform' }, 'MSForms.UserForm', undefined, undefined],
		['Design.accdb', { moduleKind: 'class', moduleType: 'accessform', designerClass: 'Access.Form' }, 'Access.Form', 'Access.Form', 'Caller'],
		['App.vbp', { moduleKind: 'userform', moduleType: 'userform', designerClass: 'VB.Form' }, undefined, undefined, undefined],
		['Doc.docm', { moduleKind: 'document', moduleType: 'document', documentType: 'document' }, undefined, 'Word.Document', undefined],
		['Book.xlsm', { moduleKind: 'document', moduleType: 'document', documentType: 'worksheet' }, undefined, 'Excel.Worksheet', undefined],
	])('keeps designer and document receiver contexts for %s', async (projectPath, metadata, meType, meHostType, meProjectType) => {
		const fixture = projectFixture();
		fixture.moduleMetadata.set('caller', { moduleName: 'Caller', ...metadata } as never);
		const doc = documentFor('Sub Demo()\nMe.Hide\nEnd Sub\n', 'file', projectPath);
		(vscode.workspace as { textDocuments: readonly vscode.TextDocument[] }).textDocuments = [doc];
		const formPaint = vi.spyOn(analyzer, 'collectImplicitMemberMethodTokens');
		const hostPaint = vi.spyOn(analyzer, 'collectHostMemberMethodTokens');
		const provider = new VbaTypeSemanticTokensProvider({ contextForProject: async () => fixture } as never);
		disposables.push(provider);
		await provider.provideDocumentSemanticTokens(doc, active);
		expect(formPaint.mock.calls[0][1]?.meType).toBe(meType);
		expect(hostPaint.mock.calls[0][1]?.meType).toBe(meHostType);
		expect(hostPaint.mock.calls[0][1]?.meProjectType).toBe(meProjectType);
	});

	it('drops a request superseded during a project-context fetch', async () => {
		const fixture = projectFixture();
		const doc = documentFor(fixture.modules[0].source, 'file');
		(vscode.workspace as { textDocuments: readonly vscode.TextDocument[] }).textDocuments = [doc];
		const service = { contextForProject: vi.fn(async () => { (doc as { version: number }).version++; return fixture; }) };
		vi.mocked(liveProjectIndexForDocument).mockResolvedValue(fixture.project);
		const provider = new VbaTypeSemanticTokensProvider(service as never);
		disposables.push(provider);
		expect((await provider.provideDocumentSemanticTokens(doc, active)).data.length).toBe(0);
		await provider.provideDocumentSemanticTokens(doc, active);
		expect(service.contextForProject).toHaveBeenCalledTimes(2);
	});
});

describe('signature help work', () => {
	it.each(['x = 1 ', 'Dim value As Long ', "' comment ", 'x = "text" '])(
		'does not load project context without an active call: %s', async line => {
			const doc = documentFor('Sub Demo()\n' + line + '\nEnd Sub\n');
			const context = {
				cachedEditorProjectContext: vi.fn(() => undefined), cheapEditorProjectContext: vi.fn(() => ({})),
				localEditorProjectContext: vi.fn(() => ({})), warmEditorProjectContext: vi.fn(),
				buildEditorProjectContextWithin: vi.fn(async () => undefined),
			};
			const provider = new VbaHoverSignatureProvider(context as never);
			expect(await provider.provideSignatureHelp(doc, new vscode.Position(1, line.length), active)).toBeUndefined();
			for (const method of Object.values(context)) expect(method).not.toHaveBeenCalled();
		});

	it('keeps active runtime call tips', async () => {
		const line = 'x = Left("hello", ';
		const doc = documentFor('Sub Demo()\n' + line + '\nEnd Sub\n');
		const provider = new VbaHoverSignatureProvider({ cachedEditorProjectContext: () => ({}) } as never);
		const help = await provider.provideSignatureHelp(doc, new vscode.Position(1, line.length), active);
		expect(help?.signatures[0].label).toContain('Left');
		expect(help?.activeParameter).toBe(1);
	});
});

describe('document highlights work', () => {
	it('searches only this document while retaining external symbol binding', async () => {
		const fixture = projectFixture(40);
		const doc = documentFor(fixture.modules[0].source);
		const spy = vi.spyOn(references, 'collectSymbolReferences');
		const provider = new VbaDocumentHighlightProvider({ contextForProject: async () => fixture } as never);
		const highlights = await provider.provideDocumentHighlights(doc, new vscode.Position(1, 1), active);
		expect(highlights).toHaveLength(2);
		expect(highlights?.map(item => item.kind)).toEqual([vscode.DocumentHighlightKind.Write, vscode.DocumentHighlightKind.Read]);
		expect(spy.mock.calls[0][0].size).toBe(1);
		expect(spy.mock.calls[0][0].get('caller')).toBe(fixture.modules[0]);
		expect(spy.mock.calls[0][2]).toBe(fixture.modules);
	});

	it('matches project-wide reference filtering for members, shadows, and code names', () => {
		const modules = [
			{ moduleName: 'Caller', type: 'standard', source: 'Sub Demo()\nDim obj As Widget\nobj.Value = 1\nDebug.Print obj.Value\nApi.Counter = Counter + 1\nSheet1.Run\nEnd Sub\n' },
			{ moduleName: 'Widget', type: 'class', source: 'Public Value As Long\n' },
			{ moduleName: 'Api', type: 'standard', source: 'Public Counter As Long\n' },
			{ moduleName: 'Sheet1', type: 'document', documentType: 'worksheet' as const, source: 'Public Sub Run()\nEnd Sub\n' },
			{ moduleName: 'Shadow', type: 'standard', source: 'Sub Other()\nDim Counter As Long\nCounter = 1\nEnd Sub\n' },
		];
		const project = buildLiveVbaProjectIndex(modules);
		const byModule = new Map(modules.map(mod => [mod.moduleName.toLowerCase(), mod]));
		const current = modules[0];
		for (const word of ['Value', 'Counter', 'Run']) {
			const offset = current.source.indexOf(word);
			const collect = (map: typeof byModule) => references.collectSymbolReferences(
				map, project, modules, current.source, current.moduleName, current, word, offset + word.length, offset + 1, true,
			).references;
			expect(collect(new Map([['caller', current]]))).toEqual(collect(byModule).filter(span => span.moduleName === 'Caller'));
		}
	});
});

it.skipIf(!process.env.XLIDE_SURFACE_BENCHMARK_OUTPUT)('measures highlighting and caret surface work', async () => {
    const { writeFileSync } = await import('node:fs');
    const results: Record<string, number> = {};
    const measure = async (name: string, work: () => unknown) => {
        await work();
        const times: number[] = [];
        for (let i = 0; i < 21; i++) {
            const start = performance.now();
            await work();
            times.push(performance.now() - start);
        }
        times.sort((a, b) => a - b);
        results[name] = Number(times[10].toFixed(4));
    };
    const fixture = projectFixture(100);
    const doc = documentFor(fixture.modules[0].source);
    const provider = new VbaDocumentHighlightProvider({ contextForProject: async () => fixture } as never);
    await measure('document highlights, 102-module project (ms)', () =>
        provider.provideDocumentHighlights(doc, new vscode.Position(1, 1), active));

    const source = Array.from({ length: 1200 }, (_, i) =>
        `Sub P${i}()\nDim n As Long\nn = 1\nEnd Sub\n`).join('');
    const lines = source.split('\n');
    const large = documentFor(source);
    vi.mocked(large.getText).mockImplementation(() => lines.join('\n'));
    const editor = { document: large, selection: { active: new vscode.Position(lines.length - 3, 5) } };
    (vscode.window as unknown as { activeTextEditor: unknown }).activeTextEditor = editor;
    const tracker = new VbaCaretProcedureTracker();
    disposables.push(tracker);
    const onSelection = vi.mocked(vscode.window.onDidChangeTextEditorSelection).mock.calls.at(-1)![0];
    const onChange = vi.mocked(vscode.workspace.onDidChangeTextDocument).mock.calls.at(-1)?.[0];
    await measure('caret tracker during body typing, 1200 procedures (ms)', () => {
        (large as { version: number }).version++;
        onChange?.({ document: large, contentChanges: [{
            text: '1', range: { start: { line: editor.selection.active.line }, end: { line: editor.selection.active.line } },
        }] } as never);
        onSelection({ textEditor: editor } as never);
    });

    const semanticDocument = documentFor(source);
    (vscode.workspace as { textDocuments: readonly vscode.TextDocument[] }).textDocuments = [semanticDocument];
    vi.mocked(semanticDocument.getText).mockImplementation(() => lines.join('\n'));
    const semantics = new VbaTypeSemanticTokensProvider({} as never);
    disposables.push(semantics);
    await measure('semantic token cache hit, 1200 procedures (ms)', () =>
        semantics.provideDocumentSemanticTokens(semanticDocument, active));
    writeFileSync(process.env.XLIDE_SURFACE_BENCHMARK_OUTPUT!, JSON.stringify({
        samples: 21, procedures: 1200, projectModules: fixture.modules.length, medianMs: results,
    }, null, 2));
});


describe('semantic background refresh lifetime', () => {
    it.each(['disposed', 'closed', 'edited'])('drops a refresh after its document/provider is %s', async reason => {
        const fixture = projectFixture();
        const doc = documentFor(fixture.modules[0].source);
        (vscode.workspace as { textDocuments: readonly vscode.TextDocument[] }).textDocuments = [doc];
        let finish!: (fixture: ReturnType<typeof projectFixture>) => void;
        const service = { contextForProject: vi.fn(() => new Promise<ReturnType<typeof projectFixture>>(resolve => { finish = resolve; })) };
        const types = vi.spyOn(fixture.project, 'visibleTypeNames');
        const provider = new VbaTypeSemanticTokensProvider(service as never);
        disposables.push(provider);
        const refreshed = vi.fn();
        provider.onDidChangeSemanticTokens(refreshed);
        vi.useFakeTimers();
        try {
            await provider.provideDocumentSemanticTokens(doc, active);
            await vi.advanceTimersByTimeAsync(400);
            expect(service.contextForProject).toHaveBeenCalledTimes(1);
            if (reason === 'disposed') { provider.dispose(); }
            if (reason === 'closed') {
                Object.defineProperty(doc, 'isClosed', { value: true });
                (vscode.workspace as { textDocuments: readonly vscode.TextDocument[] }).textDocuments = [];
            }
            if (reason === 'edited') { (doc as { version: number }).version++; }
            finish(fixture);
            await vi.advanceTimersByTimeAsync(0);
            expect(types).not.toHaveBeenCalled();
            expect(refreshed).not.toHaveBeenCalled();
        } finally { vi.useRealTimers(); }
    });
});


describe('failed background semantic refresh', () => {
    it('preserves cached tokens without firing another workspace repaint', async () => {
        const fixture = projectFixture();
        const doc = documentFor(fixture.modules[0].source);
        (vscode.workspace as { textDocuments: readonly vscode.TextDocument[] }).textDocuments = [doc];
        const service = { contextForProject: vi.fn(async () => fixture) };
        const provider = new VbaTypeSemanticTokensProvider(service as never);
        disposables.push(provider);
        const refreshed = vi.fn();
        provider.onDidChangeSemanticTokens(refreshed);
        vi.useFakeTimers();
        try {
            await provider.provideDocumentSemanticTokens(doc, active);
            await vi.advanceTimersByTimeAsync(400);
            expect(refreshed).toHaveBeenCalledTimes(1);
            const cached = await provider.provideDocumentSemanticTokens(doc, active);
            refreshed.mockClear();
            await vi.advanceTimersByTimeAsync(5001);
            service.contextForProject.mockRejectedValueOnce(new Error('temporarily unavailable'));
            expect(await provider.provideDocumentSemanticTokens(doc, active)).toBe(cached);
            await vi.advanceTimersByTimeAsync(400);
            expect(service.contextForProject).toHaveBeenCalledTimes(2);
            expect(refreshed).not.toHaveBeenCalled();
            expect(await provider.provideDocumentSemanticTokens(doc, active)).toBe(cached);
        } finally { vi.useRealTimers(); }
    });
});


describe('semantic reopened-document cache', () => {
    it('does not return old tokens when a reopened URI starts at version one again', async () => {
        const fixture = projectFixture();
        const original = documentFor('Sub Demo()\nDebug.Print ThisWorkbook.Name\nEnd Sub\n', 'file');
        (vscode.workspace as { textDocuments: readonly vscode.TextDocument[] }).textDocuments = [original];
        const provider = new VbaTypeSemanticTokensProvider({ contextForProject: async () => fixture } as never);
        disposables.push(provider);
        const oldTokens = await provider.provideDocumentSemanticTokens(original, active);
        expect(oldTokens.data.length).toBeGreaterThan(0);
        (vscode.workspace as { textDocuments: readonly vscode.TextDocument[] }).textDocuments = [];
        for (const listener of [...semanticEvents.close]) { listener(original); }
        const reopened = documentFor('Sub Demo()\nDebug.Print ActiveSheet.Name\nEnd Sub\n', 'file');
        (vscode.workspace as { textDocuments: readonly vscode.TextDocument[] }).textDocuments = [reopened];
        const newTokens = await provider.provideDocumentSemanticTokens(reopened, active);
        expect(newTokens).not.toBe(oldTokens);
        expect(reopened.getText).toHaveBeenCalled();
    });
});


describe('overlapping reopened semantic refresh', () => {
    it('does not let the old completion unlock a newer pending refresh for the same URI', async () => {
        const fixture = projectFixture();
        const original = documentFor(fixture.modules[0].source);
        (vscode.workspace as { textDocuments: readonly vscode.TextDocument[] }).textDocuments = [original];
        const finish: Array<(fixture: ReturnType<typeof projectFixture>) => void> = [];
        const service = { contextForProject: vi.fn(() => new Promise<ReturnType<typeof projectFixture>>(resolve => { finish.push(resolve); })) };
        const provider = new VbaTypeSemanticTokensProvider(service as never);
        disposables.push(provider);
        vi.useFakeTimers();
        try {
            await provider.provideDocumentSemanticTokens(original, active);
            await vi.advanceTimersByTimeAsync(400);
            (vscode.workspace as { textDocuments: readonly vscode.TextDocument[] }).textDocuments = [];
            for (const listener of [...semanticEvents.close]) { listener(original); }
            const reopened = documentFor(fixture.modules[0].source);
            (vscode.workspace as { textDocuments: readonly vscode.TextDocument[] }).textDocuments = [reopened];
            await provider.provideDocumentSemanticTokens(reopened, active);
            await vi.advanceTimersByTimeAsync(400);
            expect(service.contextForProject).toHaveBeenCalledTimes(2);
            finish[0](fixture);
            await vi.advanceTimersByTimeAsync(0);
            await provider.provideDocumentSemanticTokens(reopened, active);
            await vi.advanceTimersByTimeAsync(400);
            expect(service.contextForProject).toHaveBeenCalledTimes(2);
            finish[1](fixture);
            await vi.advanceTimersByTimeAsync(0);
        } finally { vi.useRealTimers(); }
    });
});


describe('hover and call-tip document lifetimes', () => {
    it.each(['hover', 'signature'] as const)('skips source reads for a closed %s document', async kind => {
        const doc = documentFor('Sub Demo()\nRemoteCall(\nEnd Sub\n');
        (doc as unknown as { isClosed: boolean }).isClosed = true;
        const provider = new VbaHoverSignatureProvider({
            cachedEditorProjectContext: () => ({}),
        } as never);
        const request = kind === 'hover' ? provider.provideHover.bind(provider) : provider.provideSignatureHelp.bind(provider);
        await request(doc, new vscode.Position(1, 2), active);
        expect(doc.getText).not.toHaveBeenCalled();
    });

    it.each(['hover', 'signature'] as const)('does not resolve %s against a context arriving after close', async kind => {
        const doc = documentFor('Sub Demo()\nRemoteCall(\nEnd Sub\n');
        let finish!: (value: {}) => void;
        const loaded = new Promise<{}>(resolve => { finish = resolve; });
        const resolver = kind === 'hover' ? vi.spyOn(analyzer, 'resolveHover') : vi.spyOn(analyzer, 'resolveSignatureHelp');
        resolver.mockReturnValue(undefined);
        const context = {
            cachedEditorProjectContext: () => undefined, cheapEditorProjectContext: () => ({}),
            localEditorProjectContext: () => ({}), warmEditorProjectContext() {},
            buildEditorProjectContextWithin: () => loaded,
        };
        const provider = new VbaHoverSignatureProvider(context as never);
        const request = kind === 'hover' ? provider.provideHover.bind(provider) : provider.provideSignatureHelp.bind(provider);
        const result = request(doc, new vscode.Position(1, kind === 'hover' ? 2 : 11), active);
        expect(resolver).toHaveBeenCalledTimes(2);
        (doc as unknown as { isClosed: boolean }).isClosed = true;
        finish({});
        expect(await result).toBeUndefined();
        expect(resolver).toHaveBeenCalledTimes(2);
    });
});
