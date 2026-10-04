import { beforeEach, describe, expect, it, vi } from 'vitest';
import type * as vscodeTypes from 'vscode';

vi.mock('vscode', async () => {
    const { vscodeMock, Position } = await import('./helpers/vscodeMock');
    return vscodeMock({
        workspace: { onDidCloseTextDocument: vi.fn(() => ({ dispose() {} })) },
        CompletionItem: class { constructor(public label: string, public kind: number) {} },
        CompletionList: class { constructor(public items: unknown[], public isIncomplete: boolean) {} },
        CompletionItemKind: new Proxy({}, { get: () => 1 }),
        CompletionTriggerKind: { Invoke: 0, TriggerCharacter: 1, TriggerForIncompleteCompletions: 2 },
        SnippetString: class { constructor(public value: string) {} },
        Range: class {
            start: typeof Position.prototype;
            end: typeof Position.prototype;
            constructor(line: number, start: number, endLine: number, end: number) {
                this.start = new Position(line, start); this.end = new Position(endLine, end);
            }
        },
    });
});
vi.mock('../src/analyzer/call/callContext', async () => {
    const actual = await vi.importActual<typeof import('../src/analyzer/call/callContext')>('../src/analyzer/call/callContext');
    return { ...actual, callableCompletionShouldInsertParens: vi.fn(actual.callableCompletionShouldInsertParens) };
});

import * as vscode from 'vscode';
import { VbaMemberCompletionProvider } from '../src/vbaCompletionProvider';
import { callableCompletionShouldInsertParens } from '../src/analyzer/call/callContext';
import { getWordObjectModel } from '../src/analyzer/host/wordObjectModel';
import type { EditorProjectContext, VbaEditorProjectContextService } from '../src/vbaEditorProjectContext';

function prepareRequest(line: string, context: EditorProjectContext = {}, column = line.length, prelude = 'Sub Demo()\n') {
    const source = prelude + line + '\nEnd Sub';
    const lineIndex = prelude.split('\n').length - 1;
    const document = {
        uri: { scheme: 'file', path: '/Module1.bas', toString: () => 'file:/Module1.bas' },
        version: 1, languageId: 'vba',
        getText: () => source,
        offsetAt: (position: vscode.Position) => source.split('\n').slice(0, position.line).reduce((total, text) => total + text.length + 1, 0) + position.character,
        lineAt: (index: number) => ({ text: source.split('\n')[index] }),
    };
    const projectContext = {
        cachedEditorProjectContext: vi.fn(() => context),
        localEditorProjectContext: vi.fn(() => context),
        warmEditorProjectContext: vi.fn(),
        buildEditorProjectContextWithin: vi.fn(async () => context),
    };
    const provider = new VbaMemberCompletionProvider(projectContext as unknown as VbaEditorProjectContextService);
    return { projectContext, run: () => provider.provideCompletionItems(document as unknown as vscodeTypes.TextDocument, new vscode.Position(lineIndex, column)) };
}

function request(line: string, context: EditorProjectContext = {}, column = line.length, prelude = 'Sub Demo()\n') {
    return prepareRequest(line, context, column, prelude).run();
}

beforeEach(() => { vi.mocked(callableCompletionShouldInsertParens).mockClear(); });

describe('completion provider surface', () => {
    it('skips project lookups for ordinary comments while preserving directive suggestions', async () => {
        const plain = prepareRequest("' ordinary comment");
        expect((await plain.run()).items).toEqual([]);
        expect(plain.projectContext.cachedEditorProjectContext).not.toHaveBeenCalled();
        expect(plain.projectContext.localEditorProjectContext).not.toHaveBeenCalled();
        expect(plain.projectContext.buildEditorProjectContextWithin).not.toHaveBeenCalled();
        const directive = await request("' @xlide-");
        expect(directive.items.map(item => item.label)).toContain('@xlide-test');
    });

    it('includes project classes alongside built-in types after As', async () => {
        const result = await request('Dim thing As ', { projectTypes: [{ name: 'Widget', kind: 'class', moduleName: 'Widget' }] });
        expect(result.items.map(item => item.label)).toContain('Widget');
    });
    it('uses the document host rather than Excel for type completion', async () => {
        const result = await request('Dim thing As Range', { host: 'word', hostModel: getWordObjectModel() });
        expect(result.items.find(item => item.label === 'Range')?.detail).toBe('Word type');
    });
    it.each(['caf\u00e9', '\u0915\u093e'])('replaces the whole Unicode member prefix %s including its suffix', async prefix => {
        const result = await request(`obj.${prefix}tail`, {
            projectClassMembers: [{ name: 'Widget', kind: 'class', moduleName: 'Widget',
                members: [{ name: `${prefix}Value`, kind: 'property', moduleName: 'Widget' }] }],
        }, 4 + prefix.length, 'Sub Demo()\nDim obj As Widget\n');
        const item = result.items.find(item => item.label === `${prefix}Value`);
        expect(item).toBeDefined();
        expect(item?.range).toEqual(new vscode.Range(2, 4, 2, 4 + prefix.length + 4));
        expect(item?.insertText).toBe(`${prefix}Value`);
    });
    it.each([
        ['Set value = Application.', 'Calculate($0)'],
        ['Application.', 'Calculate'],
    ])('preserves callable insertion for %s', async (line, expected) => {
        const result = await request(line);
        const item = result.items.find(item => item.label === 'Calculate');
        const insert = item?.insertText;
        expect(typeof insert === 'string' ? insert : insert?.value).toBe(expected);
    });
    it.each([
        ['value = Abs(-1)', 'Abs', 'value = Ab'],
        ['value = Application.Intersect(a, b)', 'Intersect', 'value = Application.Int'],
        ['Call Application.Calculate ()', 'Calculate', 'Call Application.Cal'],
    ])('preserves an existing argument list in %s', async (line, name, prefix) => {
        const result = await request(line, {}, prefix.length);
        const item = result.items.find(item => item.label === name);
        expect(item).toBeDefined();
        const insert = item?.insertText;
        expect(typeof insert === 'string' ? insert : insert?.value).toBe(name);
    });
    it.each([
        ['obj.[Unit Pr', 'obj.[Unit Pr'],
        ['obj.[Unit Price]', 'obj.[Unit Pr'],
        ['obj.[Unit Price]', 'obj.[Unit Price]'],
    ])('completes bracketed member names in %s at %s', async (line, prefix) => {
        const result = await request(line, {
            projectClassMembers: [{ name: 'Widget', kind: 'class', moduleName: 'Widget',
                members: [{ name: 'Unit Price', kind: 'property', moduleName: 'Widget' }] }],
        }, prefix.length, 'Sub Demo()\nDim obj As Widget\n');
        const item = result.items.find(item => item.label === 'Unit Price');
        expect(item).toBeDefined();
        expect(item?.range).toEqual(new vscode.Range(2, 4, 2, line.length));
        expect(item?.insertText).toBe('[Unit Price]');
        expect(item?.filterText).toBe('[Unit Price]');
    });
    it.each(['value = Lef$("abc", 1)', 'value = Left$("abc", 1)'])(
        'replaces the complete suffixed runtime name in %s', async line => {
            const column = line.indexOf('$');
            const result = await request(line, {}, column);
            const item = result.items.find(item => item.label === 'Left$');
            expect(item).toBeDefined();
            expect(item?.range).toEqual(new vscode.Range(1, 8, 1, column + 1));
            expect(item?.insertText).toBe('Left$');
        });
    it('does not offer unrelated globals for an unmatched bracketed member', async () => {
        expect((await request('ThisWorkbook.Sheets(1).[Cez]')).items).toEqual([]);
    });
    it('keeps parentheses insertion when a fresh dollar-suffixed function is completed', async () => {
        const result = await request('value = Lef$');
        const item = result.items.find(item => item.label === 'Left$');
        expect(item?.range).toEqual(new vscode.Range(1, 8, 1, 12));
        expect(typeof item?.insertText === 'string' ? item.insertText : item?.insertText?.value).toBe('Left$($0)');
    });
    it.each(['Err', 'Debug', 'UserForms'])('inserts the runtime object %s without call parentheses', async name => {
        const line = `Set obj = ${name.slice(0, -1)}`;
        const result = await request(line);
        const item = result.items.find(item => item.label === name);
        expect(item).toBeDefined();
        expect(item?.insertText).toBe(name);
        if (name !== 'Err') { expect(callableCompletionShouldInsertParens).not.toHaveBeenCalled(); }
    });
    it.each([
        ['Set obj = Uni', 'Union', 'Union($0)'],
        ['Call Uni', 'Union', 'Union($0)'],
        ['Uni', 'Union', 'Union'],
    ])('preserves callable host globals for %s', async (line, name, expected) => {
        const item = (await request(line)).items.find(item => item.label === name);
        expect(item).toBeDefined();
        expect(typeof item?.insertText === 'string' ? item.insertText : item?.insertText?.value).toBe(expected);
    });
    it('skips callable classification for a property-only list', async () => {
        const result = await request('ThisWorkbook.Sheets(1).ce');
        expect(result.items.map(item => item.label)).toContain('Cells');
        expect(callableCompletionShouldInsertParens).not.toHaveBeenCalled();
    });
    it('classifies callable insertion only once for a large member list', async () => {
        const result = await request('Set value = Application.');
        expect(result.items.length).toBeGreaterThan(100);
        expect(callableCompletionShouldInsertParens).toHaveBeenCalledTimes(1);
    });
});


describe('ordinary string completion work', () => {
    it.each(['value = "ordinary', 'obj.Caption = "ordinary', 'obj.Configure caption:="ordinary', 'Application.Run "Main.Go"'])('skips project context work at %s', async line => {
        const request = prepareRequest(line);
        expect((await request.run()).items).toEqual([]);
        expect(request.projectContext.cachedEditorProjectContext).not.toHaveBeenCalled();
        expect(request.projectContext.localEditorProjectContext).not.toHaveBeenCalled();
        expect(request.projectContext.warmEditorProjectContext).not.toHaveBeenCalled();
    });
});
