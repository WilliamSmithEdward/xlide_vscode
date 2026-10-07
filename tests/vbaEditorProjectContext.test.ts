import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type * as vscodeTypes from 'vscode';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
    workspace: { onDidCloseTextDocument: vi.fn(() => ({ dispose() {} })), textDocuments: [] },
}));
vi.mock('../src/vbaProjectAnalysis', async () => {
    const actual = await vi.importActual<typeof import('../src/vbaProjectAnalysis')>('../src/vbaProjectAnalysis');
    return { ...actual, buildLiveVbaProjectIndex: vi.fn(actual.buildLiveVbaProjectIndex), buildLiveVbaProjectIndexAsync: vi.fn(actual.buildLiveVbaProjectIndexAsync) };
});

import * as vscode from 'vscode';
import { VbaEditorProjectContextService } from '../src/vbaEditorProjectContext';
import { buildLiveVbaProjectIndex, buildLiveVbaProjectIndexAsync } from '../src/vbaProjectAnalysis';
import type { VbaProjectIndexService } from '../src/vbaProjectIndexService';

function project(name: string) {
    return buildLiveVbaProjectIndex([{ moduleName: 'Module', moduleKind: 'standard',
        source: `Public Type ${name}\nValue As Long\nEnd Type\n` }]);
}
function pending<T>() {
    let resolve!: (value: T) => void;
    const promise = new Promise<T>(done => { resolve = done; });
    return { promise, resolve };
}
function setup() {
    const doc = { version: 1, isClosed: false,
        uri: { scheme: 'file', path: '/Module.bas', fsPath: 'C:/loose/Module.bas', toString: () => 'file:/Module.bas' },
        getText: () => '' } as unknown as vscodeTypes.TextDocument;
    (vscode.workspace as unknown as { textDocuments: vscodeTypes.TextDocument[] }).textDocuments = [doc];
    const service = new VbaEditorProjectContextService({} as VbaProjectIndexService);
    return { doc, service };
}

beforeEach(() => { vi.mocked(buildLiveVbaProjectIndexAsync).mockReset(); vi.mocked(buildLiveVbaProjectIndex).mockClear(); });
afterEach(() => { vi.restoreAllMocks(); });

describe('editor completion context loading races', () => {
    it('keeps current-module macros separate from external bare-call procedures', () => {
        const { doc, service } = setup();
        const ctx = service.localEditorProjectContext(doc, 'Public Sub Clicked()\nEnd Sub');
        expect(ctx.macroProcedures?.map(procedure => procedure.name)).toContain('Clicked');
        expect(ctx.projectProcedures?.map(procedure => procedure.name)).not.toContain('Clicked');
    });
    it('expires reused local contexts after the cache lifetime', () => {
        const { doc, service } = setup();
        const now = vi.spyOn(Date, 'now').mockReturnValue(0);
        service.localEditorProjectContext(doc, 'Sub Demo()\nEnd Sub');
        now.mockReturnValue(11000);
        service.localEditorProjectContext(doc, 'Sub Demo()\nEnd Sub');
        expect(buildLiveVbaProjectIndex).toHaveBeenCalledTimes(2);
    });

    it('reuses local symbols for repeated requests, rebuilding for edits and invalidation', () => {
        const { doc, service } = setup();
        const source = 'Sub Demo()\nEnd Sub\n';
        const first = service.localEditorProjectContext(doc, source);
        expect(service.localEditorProjectContext(doc, source)).toBe(first);
        expect(buildLiveVbaProjectIndex).toHaveBeenCalledTimes(1);
        (doc as unknown as { version: number }).version++;
        service.localEditorProjectContext(doc, source);
        expect(buildLiveVbaProjectIndex).toHaveBeenCalledTimes(2);
        service.invalidate();
        service.localEditorProjectContext(doc, source);
        expect(buildLiveVbaProjectIndex).toHaveBeenCalledTimes(3);
    });
    it('does not reuse local symbols for a different source at the same version', () => {
        const { doc, service } = setup();
        service.localEditorProjectContext(doc, 'Public Type First\nValue As Long\nEnd Type');
        const second = service.localEditorProjectContext(doc, 'Public Type Second\nValue As Long\nEnd Type');
        expect(second.projectTypes?.map(type => type.name)).toContain('Second');
        expect(second.projectTypes?.map(type => type.name)).not.toContain('First');
        expect(buildLiveVbaProjectIndex).toHaveBeenCalledTimes(2);
    });

    it('deduplicates repeated warm requests for the same document version', async () => {
        const { doc, service } = setup();
        const load = pending<ReturnType<typeof project>>();
        vi.mocked(buildLiveVbaProjectIndexAsync).mockReturnValue(load.promise);
        service.warmEditorProjectContext(doc, 'same');
        service.warmEditorProjectContext(doc, 'same');
        const request = service.buildEditorProjectContextWithin(doc, 'same', 1000);
        expect(buildLiveVbaProjectIndexAsync).toHaveBeenCalledTimes(1);
        load.resolve(project('CurrentType'));
        expect((await request)?.projectTypes?.map(type => type.name)).toContain('CurrentType');
    });
    it('rejects a completed load for a document closed while loading', async () => {
        const { doc, service } = setup();
        const load = pending<ReturnType<typeof project>>();
        vi.mocked(buildLiveVbaProjectIndexAsync).mockReturnValue(load.promise);
        const request = service.buildEditorProjectContextWithin(doc, '', 1000);
        (doc as unknown as { isClosed: boolean }).isClosed = true;
        load.resolve(project('OldType'));
        expect((await request)?.projectTypes ?? []).toEqual([]);
        expect(service.cachedEditorProjectContext(doc)).toBeUndefined();
    });
    it('preserves the newer cache when an invalidated load finishes last', async () => {
        const { doc, service } = setup();
        const oldLoad = pending<ReturnType<typeof project>>();
        const newLoad = pending<ReturnType<typeof project>>();
        vi.mocked(buildLiveVbaProjectIndexAsync).mockReturnValueOnce(oldLoad.promise).mockReturnValueOnce(newLoad.promise);
        const oldRequest = service.buildEditorProjectContextWithin(doc, '', 1000);
        service.invalidate();
        const newRequest = service.buildEditorProjectContextWithin(doc, '', 1000);
        newLoad.resolve(project('NewType'));
        await newRequest;
        oldLoad.resolve(project('OldType'));
        await oldRequest;
        expect(service.cachedEditorProjectContext(doc)?.projectTypes?.map(type => type.name)).toContain('NewType');
        expect(service.cachedEditorProjectContext(doc)?.projectTypes?.map(type => type.name)).not.toContain('OldType');
    });

    it('does not cache results from a build invalidated while loading', async () => {
        const { doc, service } = setup();
        const load = pending<ReturnType<typeof project>>();
        vi.mocked(buildLiveVbaProjectIndexAsync).mockReturnValue(load.promise);
        const result = service.buildEditorProjectContextWithin(doc, '', 1000);
        service.invalidate();
        load.resolve(project('OldType'));
        await result;
        expect(service.cachedEditorProjectContext(doc)).toBeUndefined();
    });
    it('does not let an invalidated build answer over a newer build at the same version', async () => {
        const { doc, service } = setup();
        const oldLoad = pending<ReturnType<typeof project>>();
        const newLoad = pending<ReturnType<typeof project>>();
        vi.mocked(buildLiveVbaProjectIndexAsync).mockReturnValueOnce(oldLoad.promise).mockReturnValueOnce(newLoad.promise);
        const oldRequest = service.buildEditorProjectContextWithin(doc, '', 1000);
        service.invalidate();
        const newRequest = service.buildEditorProjectContextWithin(doc, '', 1000);
        oldLoad.resolve(project('OldType'));
        await oldRequest;
        expect(service.cachedEditorProjectContext(doc)).toBeUndefined();
        newLoad.resolve(project('NewType'));
        await newRequest;
        expect(service.cachedEditorProjectContext(doc)?.projectTypes?.map(type => type.name)).toContain('NewType');
    });
    it('warms the latest document version even while a previous build is pending', async () => {
        const { doc, service } = setup();
        const oldLoad = pending<ReturnType<typeof project>>();
        const newLoad = pending<ReturnType<typeof project>>();
        vi.mocked(buildLiveVbaProjectIndexAsync).mockReturnValueOnce(oldLoad.promise).mockReturnValueOnce(newLoad.promise);
        service.warmEditorProjectContext(doc, 'old');
        (doc as unknown as { version: number }).version++;
        service.warmEditorProjectContext(doc, 'new');
        expect(buildLiveVbaProjectIndexAsync).toHaveBeenCalledTimes(2);
        oldLoad.resolve(project('OldType')); newLoad.resolve(project('NewType'));
        await service.buildEditorProjectContextWithin(doc, '', 1000);
        expect(service.cachedEditorProjectContext(doc)?.projectTypes?.map(type => type.name)).toContain('NewType');
    });
    it('does not cache an obsolete document version when no new request starts', async () => {
        const { doc, service } = setup();
        const load = pending<ReturnType<typeof project>>();
        vi.mocked(buildLiveVbaProjectIndexAsync).mockReturnValue(load.promise);
        const request = service.buildEditorProjectContextWithin(doc, '', 1000);
        (doc as unknown as { version: number }).version++;
        load.resolve(project('OldType'));
        expect((await request)?.projectTypes ?? []).toEqual([]);
    });
});
