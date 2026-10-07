import { afterEach, describe, expect, it, vi } from 'vitest';
vi.mock('vscode', async () => ({
    ...(await import('./helpers/vscodeMock')).vscodeMock(),
    Range: class {
        start: { line: number; character: number };
        end: { line: number; character: number };
        constructor(line: number, start: number, endLine: number, end: number) {
            this.start = { line, character: start };
            this.end = { line: endLine, character: end };
        }
    },
    CompletionItem: class { constructor(public label: string, public kind: number) {} },
    CompletionList: class { constructor(public items: unknown[], public isIncomplete: boolean) {} },
    SnippetString: class { constructor(public value: string) {} },
    CompletionItemKind: { Method: 0, Property: 1, Event: 2, Function: 3, Variable: 4, Keyword: 5 },
    CompletionTriggerKind: { Invoke: 0, TriggerCharacter: 1, TriggerForIncompleteCompletions: 2 },
}));
import * as vscode from 'vscode';
import * as projectAnalysis from '../src/vbaProjectAnalysis';
import { VbaMemberCompletionProvider } from '../src/vbaCompletionProvider';
import { VbaEditorProjectContextService } from '../src/vbaEditorProjectContext';

const source = 'Sub Prompt()\nDim InstantValue As Long\ninstant\nEnd Sub';
function documentFor(text = source) {
    const lines = text.split('\n');
    const document = {
        version: 1, isClosed: false, uri: vscode.Uri.file('/instant.bas'), fileName: '/instant.bas',
        languageId: 'vba', lineCount: lines.length, getText: () => text,
        lineAt: (line: number) => ({ text: lines[line] }),
        offsetAt: (position: vscode.Position) => lines.slice(0, position.line).reduce((n, line) => n + line.length + 1, 0) + position.character,
    } as unknown as vscode.TextDocument;
    (vscode.workspace as unknown as { textDocuments: vscode.TextDocument[] }).textDocuments = [document];
    return document;
}
afterEach(() => { vi.useRealTimers(); vi.restoreAllMocks(); });

describe('completion response scheduling', () => {
    it('returns local results without waiting for a 150 ms project build', async () => {
        vi.useFakeTimers();
        const build = vi.fn(() => new Promise<undefined>(resolve => setTimeout(() => resolve(undefined), 150)));
        const warm = vi.fn();
        const context = { cachedEditorProjectContext: () => undefined, localEditorProjectContext: () => ({}),
            warmEditorProjectContext: warm, buildEditorProjectContextWithin: build } as unknown as VbaEditorProjectContextService;
        let list: vscode.CompletionList | undefined;
        void new VbaMemberCompletionProvider(context).provideCompletionItems(documentFor(), new vscode.Position(2, 7)).then(value => { list = value; });
        await vi.advanceTimersByTimeAsync(0);
        expect(list?.items.some(item => item.label === 'InstantValue')).toBe(true);
        expect(list?.isIncomplete).toBe(true);
        expect(build).not.toHaveBeenCalled();
        expect(warm).toHaveBeenCalledTimes(1);
    });

    it('picks up warmed cross-module results on the next incomplete-list request', async () => {
        let cached: Parameters<typeof projectAnalysis.projectEditorSymbolContextForModule>[0] | undefined;
        const project = projectAnalysis.buildLiveVbaProjectIndex([
            { moduleName: 'Caller', moduleKind: 'standard', source },
            { moduleName: 'Api', moduleKind: 'standard', source: 'Public Sub InstantApi()\nEnd Sub' },
        ]);
        const facts = projectAnalysis.projectEditorSymbolContextForModule(project, 'Caller');
        const context = {
            cachedEditorProjectContext: () => cached ? { moduleName: 'Caller', projectProcedures: facts.externalProjectProcedures } : undefined,
            localEditorProjectContext: () => ({ moduleName: 'Caller' }), warmEditorProjectContext: vi.fn(),
            buildEditorProjectContextWithin: vi.fn(),
        } as unknown as VbaEditorProjectContextService;
        const provider = new VbaMemberCompletionProvider(context);
        const document = documentFor();
        const first = await provider.provideCompletionItems(document, new vscode.Position(2, 7));
        expect(first.isIncomplete).toBe(true);
        expect(first.items.some(item => item.label === 'InstantApi')).toBe(false);
        cached = project;
        const next = await provider.provideCompletionItems(document, new vscode.Position(2, 7), undefined,
            { triggerKind: vscode.CompletionTriggerKind.TriggerForIncompleteCompletions });
        expect(next.items.some(item => item.label === 'InstantApi')).toBe(true);
        expect(next.isIncomplete).toBe(false);
    });

    it('starts warming after the response turn and deduplicates scheduled work', async () => {
        vi.useFakeTimers();
        const document = documentFor();
        const service = new VbaEditorProjectContextService({} as never);
        const spy = vi.spyOn(projectAnalysis, 'buildLiveVbaProjectIndexAsync');
        service.warmEditorProjectContext(document, source);
        service.warmEditorProjectContext(document, source);
        expect(spy).not.toHaveBeenCalled();
        await vi.advanceTimersByTimeAsync(0);
        expect(spy).toHaveBeenCalledTimes(1);
        expect(service.cachedEditorProjectContext(document)).toBeDefined();
    });

    it('drops scheduled warming after edits, close or invalidation', async () => {
        vi.useFakeTimers();
        const service = new VbaEditorProjectContextService({} as never);
        const spy = vi.spyOn(projectAnalysis, 'buildLiveVbaProjectIndexAsync');
        for (const action of ['edit', 'close', 'invalidate'] as const) {
            const document = documentFor();
            service.warmEditorProjectContext(document, source);
            if (action === 'edit') { (document as { version: number }).version++; }
            if (action === 'close') { (document as { isClosed: boolean }).isClosed = true; }
            if (action === 'invalidate') { service.invalidate(); }
            await vi.advanceTimersByTimeAsync(0);
        }
        expect(spy).not.toHaveBeenCalled();
    });
});
