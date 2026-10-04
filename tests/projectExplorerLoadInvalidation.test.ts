import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

const host = vi.hoisted(() => ({ findFiles: vi.fn(), showErrorMessage: vi.fn() }));
vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
    workspace: { findFiles: host.findFiles, workspaceFolders: [{ uri: { fsPath: 'C:/work' } }] },
    window: { showErrorMessage: host.showErrorMessage },
}));
import { ProjectExplorer } from '../src/projectExplorer';

const BOOK = 'C:/work/App.vbp';
type Sub = { name: string; kind: string; line: number };
type Module = { name: string; type: string; folder?: string };
let explorers: ProjectExplorer[] = [];
function deferred<T>() {
    let resolve!: (value: T) => void;
    let reject!: (reason: Error) => void;
    const promise = new Promise<T>((yes, no) => { resolve = yes; reject = no; });
    return { promise, resolve, reject };
}
function create(listModules = vi.fn(async (): Promise<Module[]> => [{ name: 'M', type: 'standard' }]),
    listSubs = vi.fn(async (): Promise<Sub[]> => [])) {
    const explorer = new ProjectExplorer({ call: vi.fn((method: string) =>
        method === 'listModules' ? listModules() : method === 'listSubs' ? listSubs()
            : Promise.resolve({ isPasswordProtected: false, isSigned: false }),
    ) } as unknown as ConstructorParameters<typeof ProjectExplorer>[0]);
    explorers.push(explorer);
    return explorer;
}
beforeEach(() => {
    explorers = [];
    host.showErrorMessage.mockClear();
    host.findFiles.mockResolvedValue([{ scheme: 'file', fsPath: BOOK }]);
});
afterEach(() => explorers.forEach(explorer => explorer.dispose()));

describe('targeted tree load invalidation', () => {
    for (const order of ['old-first', 'fresh-first'] as const) {
        it(`keeps refreshed procedures when loads finish ${order}`, async () => {
            const old = deferred<Sub[]>(), fresh = deferred<Sub[]>();
            const listSubs = vi.fn().mockReturnValueOnce(old.promise).mockReturnValue(fresh.promise);
            const explorer = create(undefined, listSubs);
            const [project] = await explorer.getChildren(), [module] = await explorer.getChildren(project);
            const pending = explorer.getChildren(module);
            explorer.refreshModuleSubs(BOOK, 'M');
            const current = explorer.getChildren(module);
            if (order === 'old-first') {
                old.resolve([{ name: 'Old', kind: 'Sub', line: 1 }]);
                await new Promise(resolve => setImmediate(resolve));
                fresh.resolve([{ name: 'Fresh', kind: 'Sub', line: 9 }]);
            } else {
                fresh.resolve([{ name: 'Fresh', kind: 'Sub', line: 9 }]);
                await current;
                old.resolve([{ name: 'Old', kind: 'Sub', line: 1 }]);
            }
            const [before, after] = await Promise.all([pending, current]);
            expect(before).toEqual(after);
            expect((await explorer.getChildren(module)).map(row => row.label)).toEqual(['Sub Fresh']);
            expect(explorer.getProcedureNode(BOOK, 'M', 'Sub Fresh')?.line).toBe(9);
            expect(listSubs).toHaveBeenCalledTimes(2);
        });
    }

    it('joins a retried project listing instead of reviving the old modules', async () => {
        const old = deferred<Module[]>(), fresh = deferred<Module[]>();
        const list = vi.fn().mockReturnValueOnce(old.promise).mockReturnValue(fresh.promise);
        const explorer = create(list), [project] = await explorer.getChildren();
        const pending = explorer.getChildren(project);
        explorer.retryLoad({ kind: 'loadError', label: 'Retry', filePath: BOOK });
        const current = explorer.getChildren(project);
        old.resolve([{ name: 'Old', type: 'standard' }]);
        await new Promise(resolve => setImmediate(resolve));
        fresh.resolve([{ name: 'Fresh', type: 'standard' }]);
        const [before, after] = await Promise.all([pending, current]);
        expect(before).toEqual(after);
        expect((await explorer.getChildren(project)).map(row => row.moduleName)).toEqual(['Fresh']);
        expect(list).toHaveBeenCalledTimes(2);
    });

    it('keeps a pending procedure cache when only a folder annotation is forgotten', async () => {
        const pending = deferred<Sub[]>(), listSubs = vi.fn(() => pending.promise);
        const explorer = create(undefined, listSubs);
        const [project] = await explorer.getChildren(), [module] = await explorer.getChildren(project);
        explorer.setModuleFolder(BOOK, 'M', 'Edited');
        const loading = explorer.getChildren(module);
        explorer.forgetModuleFolder(BOOK, 'M');
        pending.resolve([{ name: 'Run', kind: 'Sub', line: 1 }]);
        await loading;
        await explorer.getChildren(module);
        expect(listSubs).toHaveBeenCalledTimes(1);
    });

    it('ignores a procedure-load failure superseded by a save', async () => {
        const old = deferred<Sub[]>();
        const listSubs = vi.fn().mockReturnValueOnce(old.promise)
            .mockResolvedValue([{ name: 'Fresh', kind: 'Sub', line: 9 }]);
        const explorer = create(undefined, listSubs);
        const [project] = await explorer.getChildren(), [module] = await explorer.getChildren(project);
        const pending = explorer.getChildren(module);
        explorer.refreshModuleSubs(BOOK, 'M');
        const fresh = await explorer.getChildren(module);
        old.reject(new Error('Old read failed'));
        expect(await pending).toEqual(fresh);
        expect(host.showErrorMessage).not.toHaveBeenCalled();
        expect(listSubs).toHaveBeenCalledTimes(2);
    });

    it('ignores a module-load failure superseded by a retry', async () => {
        const old = deferred<Module[]>();
        const list = vi.fn().mockReturnValueOnce(old.promise)
            .mockResolvedValue([{ name: 'Fresh', type: 'standard' }]);
        const explorer = create(list), [project] = await explorer.getChildren();
        const pending = explorer.getChildren(project);
        explorer.retryLoad({ kind: 'loadError', label: 'Retry', filePath: BOOK });
        const fresh = await explorer.getChildren(project);
        old.reject(new Error('Old read failed'));
        expect(await pending).toEqual(fresh);
        expect(host.showErrorMessage).not.toHaveBeenCalled();
        expect(list).toHaveBeenCalledTimes(2);
    });
});
