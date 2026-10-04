import { afterEach, describe, expect, it, vi } from 'vitest';
import * as lexer from '../src/analyzer/lexer/tokenize';
import * as callContext from '../src/analyzer/call/callContext';
import * as projectAnalysis from '../src/vbaProjectAnalysis';
import { collectTypeNameReferences, resolveTypeReferenceAt } from '../src/analyzer/semantic/typeSemanticTokens';
import { resolveHover } from '../src/analyzer/hover/resolveHover';
import { loopIteratorSyncMayApply } from '../src/vbaSmartEnter';

vi.mock('vscode', async () => {
	const base = (await import('./helpers/vscodeMock')).vscodeMock();
	return {
		...base,
		CompletionItem: class { constructor(public label: string, public kind: number) {} },
		CompletionList: class { constructor(public items: unknown[], public isIncomplete: boolean) {} },
		SnippetString: class { constructor(public value: string) {} },
		CompletionItemKind: { Method: 0, Property: 1, Event: 2, Function: 3 },
		CompletionTriggerKind: { TriggerCharacter: 1 },
	};
});
import * as vscode from 'vscode';
import { VbaMemberCompletionProvider } from '../src/vbaCompletionProvider';
import { VbaEditorProjectContextService } from '../src/vbaEditorProjectContext';
import { registerVbaLoopIteratorSync } from '../src/vbaTypingAutomation';

function documentFor(source: string) {
	const lines = source.split('\n');
	return {
		version: 1,
		uri: vscode.Uri.file('/demo.bas'),
		fileName: '/demo.bas',
		languageId: 'vba',
		lineCount: lines.length,
		getText: vi.fn(() => source),
		lineAt: (line: number) => ({ text: lines[line] }),
		offsetAt: (pos: vscode.Position) => lines.slice(0, pos.line).reduce((n, line) => n + line.length + 1, 0) + pos.character,
	} as unknown as vscode.TextDocument;
}
afterEach(() => vi.restoreAllMocks());

describe('editor surface performance', () => {
	it.each(['ThisWorkbook.Sheets(1).', 'Call ThisWorkbook.Sheets(1).', 'Set x = ThisWorkbook.Sheets(1).'])(
		'classifies callable insertion once for all rows at %s', async expression => {
			const source = 'Sub Demo()\n' + expression + '\nEnd Sub\n';
			const document = documentFor(source);
			const ctx = { cachedEditorProjectContext: () => ({}) } as unknown as VbaEditorProjectContextService;
			const spy = vi.spyOn(callContext, 'callableCompletionShouldInsertParens');
			const list = await new VbaMemberCompletionProvider(ctx).provideCompletionItems(document, new vscode.Position(1, expression.length));
			expect(list.items.length).toBeGreaterThan(20);
			expect(list.items.some(item => item.label === 'Name')).toBe(true);
			expect(spy).toHaveBeenCalledTimes(1);
			const method = list.items.find(item => item.label === 'Activate')!;
			expect(method).toBeDefined();
			const expressionContext = expression.startsWith('Call ') || expression.startsWith('Set ');
			expect(method.insertText).toEqual(expressionContext ? { value: 'Activate($0)' } : 'Activate');
			expect(list.isIncomplete).toBe(false);
		});

	it('reuses local project facts across requests and discards them after edits or invalidation', () => {
		const source = 'Public Type Point\nx As Long\nEnd Type\nSub Demo()\nEnd Sub\n';
		const doc = documentFor(source);
		const service = new VbaEditorProjectContextService({} as never);
		const spy = vi.spyOn(projectAnalysis, 'buildLiveVbaProjectIndex');
		const first = service.localEditorProjectContext(doc, source);
		expect(service.localEditorProjectContext(doc, source)).toBe(first);
		expect(spy).toHaveBeenCalledTimes(1);
		const edited = source.replace('Point', 'Changed');
		(doc as { version: number }).version++;
		const next = service.localEditorProjectContext(doc, edited);
		expect(next).not.toBe(first);
		expect(next.projectTypes?.some(type => type.name === 'Changed')).toBe(true);
		service.invalidate();
		expect(service.localEditorProjectContext(doc, edited)).not.toBe(next);
		expect(spy).toHaveBeenCalledTimes(3);
	});

	it('does not compute diagnostic project facts for editor requests', () => {
		const project = projectAnalysis.buildLiveVbaProjectIndex([
			{ moduleName: 'Caller', source: 'Sub Demo()\nEnd Sub\n', moduleKind: 'standard' },
			{ moduleName: 'Api', source: 'Public Sub Work()\nEnd Sub\n', moduleKind: 'standard' },
		]);
		const diagnosticQueries = ['writtenNames', 'nameMentions', 'sheetChanges', 'openedFileNumbers', 'stringLiteralWords'] as const;
		const spies = diagnosticQueries.map(name => vi.spyOn(project, name));
		const context = projectAnalysis.projectEditorSymbolContextForModule(project, 'Caller');
		expect(context.externalProjectProcedures.some(proc => proc.name === 'Work')).toBe(true);
		for (const spy of spies) expect(spy).not.toHaveBeenCalled();
	});

	it('shares type spans while resolving against the current project context', () => {
		const source = 'Dim item As Widget\n';
		expect(collectTypeNameReferences(source)).toBe(collectTypeNameReferences(source));
		const offset = source.indexOf('Widget') + 1;
		expect(resolveTypeReferenceAt(source, offset)).toBeUndefined();
		expect(resolveTypeReferenceAt(source, offset, { projectTypes: [{ name: 'Widget', kind: 'class' }] })?.name).toBe('Widget');
		const changed = source.replace('Widget', 'Long');
		expect(collectTypeNameReferences(changed)).not.toBe(collectTypeNameReferences(source));
		expect(resolveTypeReferenceAt(changed, offset)?.name).toBe('Long');
		expect(resolveTypeReferenceAt(changed, -1)).toBeUndefined();
		expect(resolveTypeReferenceAt(changed, changed.length + 1)).toBeUndefined();
	});

	it('finds hovered identifiers with logarithmic token reads', () => {
		const source = 'Sub Demo()\n' + 'n = 1\n'.repeat(4000) + 'Dim value As Long\nvalue = 1\nEnd Sub\n';
		// Warm other source facts before measuring just token lookup.
		const offset = source.lastIndexOf('value');
		resolveHover(source, offset + 1);
		const original = lexer.tokenizeCached(source);
		let reads = 0;
		const indexed = new Proxy(original, { get(target, key, receiver) {
			if (typeof key === 'string' && /^\d+$/.test(key)) reads++;
			return Reflect.get(target, key, receiver);
		} });
		vi.spyOn(lexer, 'tokenizeCached').mockReturnValue(indexed);
		expect(resolveHover(source, offset + 1)?.signature).toBe('value As Long');
		expect(reads).toBeLessThan(100);
	});

	it('preserves hover behavior at token edges and whitespace', () => {
		const source = 'Sub Demo()\nDim  value As Long\nvalue=1\nEnd Sub\n';
		const start = source.lastIndexOf('value');
		for (const offset of [start, start + 1, start + 5]) {
			expect(resolveHover(source, offset)?.span).toEqual({ start, end: start + 5 });
		}
		expect(resolveHover(source, source.indexOf('Dim') + 4)).toBeUndefined();
	});

	it('skips full document reads during ordinary typing', async () => {
		const doc = documentFor('Sub Demo()\nThisWorkbook.Sheets(1).\nEnd Sub\n');
		vi.mocked(vscode.window).activeTextEditor = { document: doc } as vscode.TextEditor;
		registerVbaLoopIteratorSync({ subscriptions: [] } as unknown as vscode.ExtensionContext);
		const listener = vi.mocked(vscode.workspace.onDidChangeTextDocument).mock.calls.at(-1)![0];
		await listener({ document: doc, contentChanges: [{ text: '.', range: { start: { line: 1, character: 21 } } }] } as never);
		expect(doc.getText).not.toHaveBeenCalled();
	});
});

describe('loop iterator line gate', () => {
	it.each([
		['For index = 1 To 10', 9, true],
		['For Each item In items', 13, true],
		['Next index', 10, true],
		['  Next индекс', 12, true],
		['For index = 1 To limit', 20, false],
		['Next index, other', 10, false],
		["' Next index", 12, false],
		['Debug.Print "Next index"', 15, false],
		['ThisWorkbook.Sheets(1).', 22, false],
	])('classifies %s at column %s', (line, column, expected) => {
		expect(loopIteratorSyncMayApply(line, column)).toBe(expected);
	});
});

it.skipIf(!process.env.XLIDE_EDITOR_BENCHMARK_OUTPUT)('measures editor latency on a 1200-procedure module', async () => {
	const { writeFileSync } = await import('node:fs');
	const results: Record<string, number> = {};
	const source = 'Option Explicit\n' + Array.from({ length: 1200 }, (_, i) =>
		`Sub P${i}()\nDim n As Long\nn = 1\nEnd Sub\n`).join('') +
		'Sub Demo()\nThisWorkbook.Sheets(1).Name = "x"\nEnd Sub\n';
	const lines = source.split('\n');
	const line = lines.findIndex(text => text.startsWith('ThisWorkbook.'));
	const position = new vscode.Position(line, 'ThisWorkbook.Sheets(1).'.length);
	const dot = source.indexOf('.Name');
	const doc = documentFor(source);
	// VS Code materializes module text from its line array.
	vi.mocked(doc.getText).mockImplementation(() => lines.join('\n'));
	const service = new VbaEditorProjectContextService({} as never);
	const provider = new VbaMemberCompletionProvider(service);
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
	await measure('completion provider, repeated requests (ms)', () => provider.provideCompletionItems(doc, position));
	await measure('hover on Name, warm (ms)', () => resolveHover(source, dot + 2));
	await measure('type references, warm (ms)', () => collectTypeNameReferences(source));
	vi.mocked(vscode.window).activeTextEditor = { document: doc } as vscode.TextEditor;
	registerVbaLoopIteratorSync({ subscriptions: [] } as unknown as vscode.ExtensionContext);
	const listener = vi.mocked(vscode.workspace.onDidChangeTextDocument).mock.calls.at(-1)![0];
	const event = { document: doc, contentChanges: [{ text: '.', range: { start: { line, character: position.character - 1 } } }] };
	await measure('ordinary typing, loop synchronization listener (ms)', () => listener(event as never));
	let edit = 0;
	await measure('completion provider after a source edit (ms)', () => {
		const nextSource = source + "' edit " + edit++;
		const nextDoc = documentFor(nextSource);
		return provider.provideCompletionItems(nextDoc, position);
	});
	writeFileSync(process.env.XLIDE_EDITOR_BENCHMARK_OUTPUT!, JSON.stringify({
		procedures: 1200, sourceBytes: Buffer.byteLength(source), samples: 21, medianMs: results,
	}, null, 2));
});
