import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
const host = vi.hoisted(() => ({ findFiles: vi.fn(), showErrorMessage: vi.fn() }));
vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
    workspace: { findFiles: host.findFiles, workspaceFolders: [{ uri: { fsPath: 'C:/work' } }] },
    window: { showErrorMessage: host.showErrorMessage },
}));
import { ProjectExplorer } from '../src/projectExplorer';
const BOOK = 'C:/work/Book.xlsm';
const FILES = [{ scheme: 'file', fsPath: BOOK }];
let explorers: ProjectExplorer[] = [];
function deferred<T>() {
    let resolve!: (value: T) => void, reject!: (error: Error) => void;
    const promise = new Promise<T>((yes, no) => { resolve = yes; reject = no; });
    return { promise, resolve, reject };
}
function create(read: (method: string) => Promise<unknown> = async method =>
    method === 'listModules' ? [{ name: 'M', type: 'standard' }]
        : method === 'listWorkbookSheets' ? { sheets: [] }
            : { isPasswordProtected: false, isSigned: false }) {
    const call = vi.fn(read);
    const explorer = new ProjectExplorer({ call } as unknown as ConstructorParameters<typeof ProjectExplorer>[0]);
    explorers.push(explorer);
    return { explorer, call };
}
beforeEach(() => {
    explorers = [];
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout'] });
    host.findFiles.mockReset().mockResolvedValue(FILES);
    host.showErrorMessage.mockClear();
});
afterEach(() => { explorers.forEach(explorer => explorer.dispose()); vi.useRealTimers(); });
describe('disposed tree providers', () => {
    it('drops pending root discovery without restarting it', async () => {
        const files = deferred<typeof FILES>();
        host.findFiles.mockReturnValue(files.promise);
        const { explorer, call } = create();
        const loading = explorer.getChildren();
        explorer.dispose();
        files.resolve(FILES);
        expect(await loading).toEqual([]);
        expect(await explorer.getChildren()).toEqual([]);
        expect(host.findFiles).toHaveBeenCalledTimes(1);
        expect(call).not.toHaveBeenCalled();
    });
    for (const outcome of ['success', 'failure'] as const) {
        it(`drops a pending module read's ${outcome} without starting more work`, async () => {
            const modules = deferred<Array<{ name: string; type: string }>>();
            const { explorer, call } = create(() => modules.promise);
            const [project] = await explorer.getChildren();
            const loading = explorer.getChildren(project);
            explorer.dispose();
            if (outcome === 'success') { modules.resolve([{ name: 'Late', type: 'standard' }]); }
            else { modules.reject(new Error('Late failure')); }
            expect(await loading).toEqual([]);
            await vi.advanceTimersByTimeAsync(5000);
            expect(call.mock.calls.map(([method]) => method)).toEqual(['listModules']);
            expect(host.showErrorMessage).not.toHaveBeenCalled();
            expect(explorer.getModuleNode(BOOK, 'Late')).toBeUndefined();
        });
    }
    it('drops a pending procedure read without caching rows or scheduling a probe', async () => {
        const subs = deferred<Array<{ name: string; kind: string; line: number }>>();
        const { explorer, call } = create(async method => method === 'listModules'
            ? [{ name: 'M', type: 'standard' }] : method === 'listSubs' ? subs.promise : { sheets: [] });
        const [project] = await explorer.getChildren(), [module] = await explorer.getChildren(project);
        const loading = explorer.getChildren(module);
        const before = call.mock.calls.length;
        explorer.dispose();
        subs.resolve([{ name: 'Late', kind: 'Sub', line: 1 }]);
        expect(await loading).toEqual([]);
        expect(await explorer.getChildren(module)).toEqual([]);
        await vi.advanceTimersByTimeAsync(5000);
        expect(call).toHaveBeenCalledTimes(before);
        expect(explorer.getProcedureNode(BOOK, 'M', 'Sub Late')).toBeUndefined();
    });
    it('cancels a protection read queued by an idle timer before it reaches the bridge', async () => {
        const { explorer, call } = create();
        const [project] = await explorer.getChildren();
        await explorer.getChildren(project);
        vi.advanceTimersByTime(2000);
        explorer.dispose();
        await Promise.resolve();
        expect(call.mock.calls.some(([method]) => method === 'getProtectionInfo')).toBe(false);
    });
});
