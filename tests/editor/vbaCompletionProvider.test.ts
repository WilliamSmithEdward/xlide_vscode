import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type * as vscodeTypes from 'vscode';

vi.mock('vscode', async () => {
    const { vscodeMock, Position } = await import('../helpers/vscodeMock');
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
vi.mock('../../src/analyzer/call/callContext', async () => {
    const actual = await vi.importActual<typeof import('../src/analyzer/call/callContext')>('../../src/analyzer/call/callContext');
    return { ...actual, callableCompletionShouldInsertParens: vi.fn(actual.callableCompletionShouldInsertParens) };
});

import * as vscode from 'vscode';
import { projectOptions } from '../analyzer/diagnostics/helpers';
import { hostObjectModelForToken } from '../../src/analyzer/host/hostRegistry';
import { OFFICE_HOST_APPS, type OfficeHostApp } from '../../src/officeHostApps';
import * as lexer from '../../src/analyzer/lexer/tokenize';
import * as identifiers from '../../src/analyzer/completion/identifierCompletion';
import { VbaMemberCompletionProvider } from '../../src/vbaCompletionProvider';
import { callableCompletionShouldInsertParens } from '../../src/analyzer/call/callContext';
import { getWordObjectModel } from '../../src/analyzer/host/wordObjectModel';
import { getExcelObjectModel } from '../../src/analyzer/host/excelObjectModel';
import { getPowerPointObjectModel } from '../../src/analyzer/host/powerpointObjectModel';
import { getAccessObjectModel } from '../../src/analyzer/host/accessObjectModel';
import type { EditorProjectContext, VbaEditorProjectContextService } from '../../src/vbaEditorProjectContext';

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
        readyEditorProjectContext: vi.fn<() => EditorProjectContext | undefined>(() => undefined),
        localEditorProjectContext: vi.fn(() => context),
        warmEditorProjectContext: vi.fn(),
        buildEditorProjectContextWithin: vi.fn(async () => context),
    };
    const provider = new VbaMemberCompletionProvider(projectContext as unknown as VbaEditorProjectContextService);
    return { projectContext, run: (trigger?: string) => provider.provideCompletionItems(document as unknown as vscodeTypes.TextDocument, new vscode.Position(lineIndex, column), undefined, trigger ? {triggerKind: vscode.CompletionTriggerKind.TriggerCharacter, triggerCharacter: trigger} : undefined) };
}

function request(line: string, context: EditorProjectContext = {}, column = line.length, prelude = 'Sub Demo()\n') {
    return prepareRequest(line, context, column, prelude).run();
}

beforeEach(() => { vi.mocked(callableCompletionShouldInsertParens).mockClear(); });
afterEach(() => vi.restoreAllMocks());

describe('completion provider surface', () => {
    it('prioritizes assignment values while retaining ordinary variables', async () => {
        const result = await request('ActiveCell.HorizontalAlignment = ', {}, undefined, 'Sub Demo()\nDim chosen As XlHAlign\n');
        expect(result.items.map(item => item.label)).toContain('chosen');
        expect(result.items.find(item => item.label === 'xlHAlignLeft')?.sortText).toMatch(/^0:/);
    });
    it.each(['Dim xlHAlignLeft As Long', 'Const xlHAlignLeft = 123'])('qualifies the enum constant shadowed by %s', async declaration => {
        const result = await request('ActiveCell.HorizontalAlignment = xlH', {}, undefined, `${declaration}\nSub Demo()\n`);
        expect(result.items.find(item => item.label === 'xlHAlignLeft')?.insertText).toBe('Excel.XlHAlign.xlHAlignLeft');
        expect(result.items.filter(item => item.label === 'xlHAlignLeft')).toHaveLength(2);
    });
    it('does not let a source enum replace a same-named host constant', async () => {
        const prelude = 'Enum XlHAlign\nxlHAlignLeft = 123\nEnd Enum\nSub Demo()\n';
        const result = await request('ActiveCell.HorizontalAlignment = xlH', {}, undefined, prelude);
        expect(result.items.find(item => item.label === 'xlHAlignLeft')?.insertText).toBe('Excel.XlHAlign.xlHAlignLeft');
        expect(result.items.filter(item => item.label === 'xlHAlignLeft')).toHaveLength(2);
    });
    it.each(['If x =', 'Debug.Print x =', "' x =", 'Set item ='])('rejects automatic equals without project loading: %s', async line => {
        const prepared = prepareRequest(line);
        expect((await prepared.run('=')).items).toEqual([]);
        expect(prepared.projectContext.cachedEditorProjectContext).not.toHaveBeenCalled();
        expect(prepared.projectContext.localEditorProjectContext).not.toHaveBeenCalled();
    });
    it.each(['=', ' '])('keeps an unknown property quiet on %s', async trigger => {
        expect((await prepareRequest('ActiveCell.Value = ').run(trigger)).items).toEqual([]);
    });
    it('skips project lookups for ordinary comments while preserving directive suggestions', async () => {
        const plain = prepareRequest("' ordinary comment");
        expect((await plain.run()).items).toEqual([]);
        expect(plain.projectContext.cachedEditorProjectContext).not.toHaveBeenCalled();
        expect(plain.projectContext.localEditorProjectContext).not.toHaveBeenCalled();
        expect(plain.projectContext.buildEditorProjectContextWithin).not.toHaveBeenCalled();
        const directive = await request("' @xlide-");
        expect(directive.items.map(item => item.label)).toContain('@xlide-test');
    });
    const directiveLabels = (items: { label: unknown }[]) => items.map(item => item.label)
        .filter(label => typeof label === 'string' && label.startsWith('@'));
    const hosts: [string, EditorProjectContext][] = [
        ['Excel', { host: 'excel', hostModel: getExcelObjectModel() }],
        ['Word', { host: 'word', hostModel: getWordObjectModel() }],
        ['PowerPoint', { host: 'powerpoint', hostModel: getPowerPointObjectModel() }],
        ['Access', { host: 'access', hostModel: getAccessObjectModel() }],
    ];
    // U+0019 is whitespace to the lexer after a continuation, though not to \s.
    const eom = String.fromCharCode(0x19);
    it.each(hosts.flatMap(([name, context]) => ['Dim t _', 'x = 1 + _', 'Dim t As Long, _', 'Dim t As Long _', 'x = 1: Dim t _ \t', `Dim t As Long _${eom}`]
        .map(above => [name, above, context] as const)))('%s: offers no test directives on the row continuing %j', async (_name, above, context) => {
        const prepared = prepareRequest('    ', context, 4, `Sub Demo()\n${above}\n`);
        expect(directiveLabels((await prepared.run(' ')).items)).toEqual([]);
        expect(directiveLabels((await prepared.run()).items)).toEqual([]);
    });
    it.each([['    @xl', 'Dim t _'], ['    ', 'Dim t As Long _\u3000']])('offers no test directives at %j on the row continuing %j', async (line, above) => {
        expect(directiveLabels((await prepareRequest(line, {}, line.length, `Sub Demo()\n${above}\n`).run()).items)).toEqual([]);
    });
    it.each(['', "' Checks the total _\n", "' Checks _\n' the total _\n", "x = 1 ' note _\n", 'Rem note _\n', 'x = "text _\n', 'abc_\n'])('still offers test directives on the row after %j', async above => {
        const prepared = prepareRequest('    ', {}, 4, `Sub Demo()\n${above}`);
        const directives = ['@xlide-test', '@xlide-test-skip', '@xlide-test-xfail'];
        expect(directiveLabels((await prepared.run(' ')).items)).toEqual(directives);
        expect(directiveLabels((await prepared.run()).items)).toEqual(directives);
    });

    it('includes project classes alongside built-in types after As', async () => {
        const result = await request('Dim thing As ', { projectTypes: [{ name: 'Widget', kind: 'class', moduleName: 'Widget' }] });
        expect(result.items.map(item => item.label)).toContain('Widget');
    });
    it('uses the document host rather than Excel for type completion', async () => {
        const result = await request('Dim thing As Range', { host: 'word', hostModel: getWordObjectModel() });
        expect(result.items.find(item => item.label === 'Range')?.detail).toBe('Word type');
    });
    const hostGlobals: Record<OfficeHostApp, string> = { excel: 'ActiveSheet', word: 'ActiveDocument', powerpoint: 'ActivePresentation', access: 'DoCmd' };
    describe.each(Object.keys(OFFICE_HOST_APPS) as OfficeHostApp[])('after a word and a space in %s', host => {
        const context = { host, hostModel: hostObjectModelForToken(host) };
        const others = Object.values(hostGlobals).filter(name => name !== hostGlobals[host]);
        const prelude = 'Sub Helper()\nEnd Sub\nSub Demo()\nDim total As Long\n';
        const labels = async (line: string, trigger?: string, before = prelude) =>
            (await prepareRequest(line, context, undefined, before).run(trigger)).items.map(item => String(item.label));
        it.each(['Dim t As [Long] ', 'Dim r As [Range] ', 'Dim r As Excel.[Range] ', 'Dim r As [Excel].[Range] ', 'Set c = New [Collection] ', 'Set c = New [VBA].[Collection] '])('ends a bracketed type list at %j', async line => {
            for (const trigger of [undefined, ' ']) expect(await labels(line, trigger)).toEqual([]);
        });
        it('keeps only As after a bracketed declaration name', async () => {
            for (const trigger of [undefined, ' ']) expect(await labels('Dim [my value] ', trigger)).toEqual(['As']);
        });
        it('uses no more tokenization calls than the plain type', async () => {
            const tokenize = vi.spyOn(lexer, 'tokenizeCached');
            await labels('Set c = New Collection ');
            await labels('Set c = New [Collection] ');
            tokenize.mockClear();
            await labels('Set c = New Collection ');
            const plainCalls = tokenize.mock.calls.length;
            tokenize.mockClear();
            await labels('Set c = New [Collection] ');
            expect(tokenize.mock.calls.length).toBeLessThanOrEqual(plainCalls);
        });
        it('leaves unfinished brackets outside the finished-word gate', async () => {
            expect(await labels('Dim t As [Lo')).toContain('total');
        });
        it('preserves expressions following a bracketed value', async () => {
            expect(await labels('x = [total] + ')).toContain('total');
            expect(await labels('Call [Helper] ')).toContain('total');
        });
        it.each(['Call ', 'If ', 'MsgBox ', 'Debug.Print ', 'For i = 1 To ', 'Call total$ '])('offers the general list on Ctrl+Space after %j', async line => {
            const got = await labels(line);
            expect(got).toEqual(expect.arrayContaining(['total', 'Helper', 'MsgBox', hostGlobals[host]]));
            expect(got.filter(label => others.includes(label))).toEqual([]);
        });
        it('offers the general list on Ctrl+Space on a continued row', async () => {
            expect(await labels('    ', undefined, `${prelude}Call _\n`)).toEqual(expect.arrayContaining(['total', 'Helper', 'MsgBox']));
        });
        it.each(['Call ', 'If ', 'For i = 1 To ', 'For Each v In ', 'Not ', 'Call total$ ', 'If total% '])('opens no general list on a typed space after %j', async line => {
            expect(await labels(line, ' ')).toEqual([]);
        });
        it('opens no general list on a typed space on a continued row', async () => {
            // The test directives offered on any continued row are a separate issue.
            expect((await labels('    ', ' ', `${prelude}Call _\n`)).filter(label => !label.startsWith('@'))).toEqual([]);
        });
        it.each([
            ['Call He', undefined, 'Helper'],
            ['MsgBox to', undefined, 'total'],
            ['Debug.Pr', undefined, 'Print'],
            ['Exit ', ' ', 'Sub'],
            ['x = ', undefined, 'total'],
            // Assert's Boolean argument, as `Call Paint ` offers an enum's values.
            ['Debug.Assert ', ' ', 'True'],
        ])('still completes %j on trigger %j', async (line, trigger, expected) => {
            expect(await labels(line, trigger)).toContain(expected);
        });
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

describe('completion provider cursor work', () => {
    it.each(["' ordinary comment", 'value = "ordinary', 'Application.Run "Main.Go"'])(
        'does not copy preceding procedures to classify %s', async line => {
            const prelude = Array.from({ length: 1200 }, (_, i) =>
                'Sub Padding' + i + '()\nDebug.Print ' + i + '\nEnd Sub\n').join('') + 'Sub Probe()\n';
            const prepared = prepareRequest(line, {}, line.length, prelude);
            const tokens = lexer.tokenizeCached(prelude + line + '\nEnd Sub');
            let reads = 0;
            vi.spyOn(lexer, 'tokenizeCached').mockReturnValue(new Proxy(tokens, {
                get(target, property, receiver) {
                    if (typeof property === 'string' && /^\d+$/.test(property)) { reads++; }
                    return Reflect.get(target, property, receiver);
                },
            }));
            expect((await prepared.run()).items).toEqual([]);
            expect(reads).toBeLessThan(100);
            expect(prepared.projectContext.cachedEditorProjectContext).not.toHaveBeenCalled();
            expect(prepared.projectContext.localEditorProjectContext).not.toHaveBeenCalled();
        });

    it('preserves comment continuations when the caret is on the following physical line', async () => {
        const prepared = prepareRequest('ThisWorkbook.Sheets(1).ce', {}, undefined,
            "Sub Probe()\n' continued note _\n");
        expect((await prepared.run()).items).toEqual([]);
        expect(prepared.projectContext.cachedEditorProjectContext).not.toHaveBeenCalled();
    });
});

it('serves exported assignment values from the loaded index while a background loader cannot finish', async () => {
 const source = 'Sub Demo()\nSnapshot.Flag =\nEnd Sub';
 const options = projectOptions([{ moduleName: 'Caller', source }, { moduleName: 'Types', source: 'Public Type Record\nFlag As Boolean\nEnd Type\nPublic Property Get Snapshot() As Record\nEnd Property' }], 'Caller');
 const ready = { projectClassMembers: options.projectClassMembers, projectSymbols: options.projectVisibleSymbols };
 const prepared = prepareRequest('Snapshot.Flag =', ready);
 prepared.projectContext.cachedEditorProjectContext.mockReturnValue(undefined as never);
 prepared.projectContext.readyEditorProjectContext.mockReturnValue(ready);
 prepared.projectContext.buildEditorProjectContextWithin.mockImplementation(() => new Promise(() => {}));
 expect((await prepared.run('=')).items.slice(0, 2).map(item => item.label)).toEqual(['True', 'False']);
 expect(prepared.projectContext.localEditorProjectContext).not.toHaveBeenCalled();
 expect(prepared.projectContext.warmEditorProjectContext).not.toHaveBeenCalled();
 expect(prepared.projectContext.buildEditorProjectContextWithin).not.toHaveBeenCalled();
});
it('serves local assignment values immediately when the project is cold and its loader cannot finish', async () => {
 const prepared = prepareRequest('flag =', {}, undefined, 'Sub Demo()\nDim flag As Boolean\n');
 prepared.projectContext.cachedEditorProjectContext.mockReturnValue(undefined as never);
 prepared.projectContext.buildEditorProjectContextWithin.mockImplementation(() => new Promise(() => {}));
 expect((await prepared.run('=')).items.slice(0, 2).map(item => item.label)).toEqual(['True', 'False']);
 expect(prepared.projectContext.buildEditorProjectContextWithin).not.toHaveBeenCalled();
});

it('avoids a shadowed library name when inserting a host enum constant',async()=>{
 const result=await request('sh.Visible = msoT',{},undefined,'Sub Demo(ByVal sh As Shape)\nDim Office As Long\nDim msoTrue As Long\n');
 expect(result.items.find(item=>item.label==='msoTrue')?.insertText).toBe('MsoTriState.msoTrue');
});

it('keeps the ordinary binding when every enum qualifier is shadowed',async()=>{
 const result=await request('sh.Visible = msoT',{},undefined,'Sub Demo(ByVal sh As Shape)\nDim Office As Long\nDim MsoTriState As Long\nDim msoTrue As Long\n');
 const items=result.items.filter(item=>item.label==='msoTrue');
 expect(items).toHaveLength(1);
 expect(items[0]?.sortText ?? '').not.toMatch(/^0:/);
});

it('uses the setter enum owner when another module has the same enum and member names',async()=>{
 const prelude='Private Enum Direction\nNorth = 123\nEnd Enum\nSub Demo()\n';
 const source=prelude+'Types.State = Nor\nEnd Sub';
 const options=projectOptions([{moduleName:'Caller',source},{moduleName:'Types',source:'Public Enum Direction\nNorth = 1\nEnd Enum\nPublic Property Let State(ByVal value As Direction)\nEnd Property'}],'Caller');
 const result=await request('Types.State = Nor',{moduleName:'Caller',projectClassMembers:options.projectClassMembers,projectSymbols:options.projectVisibleSymbols},undefined,prelude);
 const choices=result.items.filter(item=>item.label==='North');
 expect(choices[0]?.insertText).toBe('Types.Direction.North');
 expect(choices).toHaveLength(2);
});

it('does not build an unfiltered binding list for ordinary enum values',async()=>{
 const resolve=vi.spyOn(identifiers,'resolveIdentifierCompletions');
 await request('sh.Visible = mso',{},undefined,'Sub Demo(ByVal sh As Shape)\n');
 expect(resolve).toHaveBeenCalledTimes(1);
});
it('reuses the unfiltered binding list across shadowed enum constants',async()=>{
 const resolve=vi.spyOn(identifiers,'resolveIdentifierCompletions');
 const result=await request('sh.Visible = mso',{},undefined,'Sub Demo(ByVal sh As Shape)\nDim Office As Long\nDim msoTrue As Long\nDim msoFalse As Long\n');
 expect(result.items.find(item=>item.label==='msoTrue')?.insertText).toBe('MsoTriState.msoTrue');
 expect(result.items.find(item=>item.label==='msoFalse')?.insertText).toBe('MsoTriState.msoFalse');
 expect(resolve).toHaveBeenCalledTimes(2);
});

/** Asks the provider at a position, with `context` as the module's project context and `options` as its editor's. */
function requestAt(source: string, line: number, column: number, context: EditorProjectContext, options?: object) {
    const lines = source.split('\n');
    const document = {
        uri: { scheme: 'xlide-vba', path: `/Book.xlsm/${context.moduleName}.bas`, toString: () => `xlide-vba:/Book.xlsm/${context.moduleName}.bas` },
        version: 1, languageId: 'vba',
        getText: () => source,
        offsetAt: (position: vscode.Position) => lines.slice(0, position.line).reduce((total, text) => total + text.length + 1, 0) + position.character,
        lineAt: (index: number) => ({ text: lines[index] }),
    };
    Object.assign(vscode.window, { visibleTextEditors: options ? [{ document, options }] : [] });
    const projectContext = {
        cachedEditorProjectContext: vi.fn(() => context),
        readyEditorProjectContext: vi.fn(() => undefined),
        localEditorProjectContext: vi.fn(() => context),
        warmEditorProjectContext: vi.fn(),
        buildEditorProjectContextWithin: vi.fn(async () => context),
    };
    const provider = new VbaMemberCompletionProvider(projectContext as unknown as VbaEditorProjectContextService);
    return provider.provideCompletionItems(document as unknown as vscodeTypes.TextDocument, new vscode.Position(line, column));
}

describe('inserted blocks in the editor indentation', () => {
    afterEach(() => { Object.assign(vscode.window, { visibleTextEditors: [] }); });

    it.each([
        [{ insertSpaces: true, tabSize: 8, indentSize: 3 }, '   '],
        [{ insertSpaces: false, tabSize: 4, indentSize: 4 }, '\t'],
        [undefined, '    '],
    ])('writes a block snippet body in the editor indentation %j', async (options, unit) => {
        const result = await requestAt('Sub Demo()\n  Wh\nEnd Sub\n', 1, 4, { moduleName: 'Module1', moduleKind: 'standard' }, options);
        const item = result.items.find(candidate => candidate.label === 'While')!;
        expect((item.insertText as vscodeTypes.SnippetString).value).toBe(`While \${1:condition}\n\n  ${unit}$0\n\n  Wend`);
        expect(item.keepWhitespace).toBe(true);
    });

    it.each([
        ['an Excel workbook module', { moduleName: 'ThisWorkbook', moduleKind: 'document', documentType: 'workbook', host: 'excel' }, 'Workbook_O', 'Workbook_Open'],
        ['a Word document module', { moduleName: 'ThisDocument', moduleKind: 'document', documentType: 'document', host: 'word' }, 'Document_O', 'Document_Open'],
        ['a PowerPoint UserForm', { moduleName: 'UserForm1', moduleKind: 'userform', host: 'powerpoint' }, 'UserForm_I', 'UserForm_Initialize'],
    ] as const)('leaves a handler stub tab in %s for VS Code to convert', async (_, context, typed, handler) => {
        const result = await requestAt(`Option Explicit\n${typed}\n`, 1, typed.length, context as EditorProjectContext, { insertSpaces: true, tabSize: 2, indentSize: 2 });
        const item = result.items.find(candidate => candidate.label === handler)!;
        expect((item.insertText as vscodeTypes.SnippetString).value).toBe(`Private Sub ${handler}()\n\t$0\nEnd Sub`);
        expect(item.keepWhitespace).toBeUndefined();
    });
});

describe('a class module offers its handlers beside the other items', () => {
    const labels = (list: vscodeTypes.CompletionList) => list.items.map(item => item.label as string);
    // The empty line right below a declaration, and before the first procedure.
    const source = 'Option Explicit\nPrivate cTotal As Long\n\nSub Helper()\nEnd Sub\n';
    afterEach(() => { Object.assign(vscode.window, { visibleTextEditors: [] }); });

    it.each(['excel', 'word', 'powerpoint', 'access'] as const)('offers a %s class module its handlers beside the keywords and identifiers', async host => {
        const context: EditorProjectContext = { moduleName: 'Account', moduleKind: 'class', host };
        const result = await requestAt(source, 2, 0, context);
        expect(labels(result)).toEqual(expect.arrayContaining(['Class_Initialize', 'Class_Terminate', 'Private Sub', 'Dim']));
        const typed = await requestAt(source.replace('Long\n\n', 'Long\nC\n'), 2, 1, context);
        expect(labels(typed)).toEqual(expect.arrayContaining(['Class_Initialize', 'Class_Terminate', 'Call', 'cTotal']));
        const item = result.items.find(candidate => candidate.label === 'Class_Initialize')!;
        expect((item.insertText as vscodeTypes.SnippetString).value).toBe('Private Sub Class_Initialize()\n\t$0\nEnd Sub');
        expect(item.keepWhitespace).toBeUndefined();
    });

    it.each([
        ['an Excel workbook module', { moduleName: 'ThisWorkbook', moduleKind: 'document', documentType: 'workbook', host: 'excel' }, 'Workbook_Open'],
        ['a Word document module', { moduleName: 'ThisDocument', moduleKind: 'document', documentType: 'document', host: 'word' }, 'Document_Open'],
        ['a PowerPoint UserForm', { moduleName: 'UserForm1', moduleKind: 'userform', host: 'powerpoint' }, 'UserForm_Initialize'],
    ] as const)('offers %s its own handlers in place of the list, and no Class_ handler', async (_, context, own) => {
        const offered = labels(await requestAt(source, 2, 0, context as EditorProjectContext));
        expect(offered).toContain(own);
        expect(offered.filter(label => /^Class_/.test(label) || label === 'Dim')).toEqual([]);
    });

    it('offers an Access form module no Class_ handler', async () => {
        const offered = labels(await requestAt(source, 2, 0, { moduleName: 'Form_Orders', moduleKind: 'userform', host: 'access', meType: 'Access.Form' }));
        expect(offered.filter(label => /^Class_/.test(label))).toEqual([]);
    });
});
