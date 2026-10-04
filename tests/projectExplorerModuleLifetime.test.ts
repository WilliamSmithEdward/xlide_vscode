import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

const host = vi.hoisted(() => ({ findFiles: vi.fn() }));
vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
    workspace: { findFiles: host.findFiles, workspaceFolders: [{ uri: { fsPath: 'C:/work' } }] },
}));
import { ProjectExplorer } from '../src/projectExplorer';

const BOOK = 'C:/work/App.vbp';
type Module = { name: string; type: string; folder?: string; documentType?: string; filePath?: string };
let explorers: ProjectExplorer[] = [];
function create(list: () => Promise<Module[]>, subs = vi.fn(async () => [{ name: 'Run', kind: 'Sub', line: 1 }])) {
    const explorer = new ProjectExplorer({ call: vi.fn((method: string) =>
        method === 'listModules' ? list() : method === 'listSubs' ? subs()
            : Promise.resolve({ isPasswordProtected: false, isSigned: false }),
    ) } as unknown as ConstructorParameters<typeof ProjectExplorer>[0]);
    explorers.push(explorer);
    return explorer;
}
beforeEach(() => {
    explorers = [];
    host.findFiles.mockResolvedValue([{ scheme: 'file', fsPath: BOOK }]);
});
afterEach(() => explorers.forEach(explorer => explorer.dispose()));

describe('module rows across targeted project relisting', () => {
    it('refreshes git badges without reading 1000 modules in another project', async () => {
        const other = 'C:/work/Other.vbp';
        host.findFiles.mockResolvedValue([{ scheme: 'file', fsPath: BOOK }, { scheme: 'file', fsPath: other }]);
        const list = vi.fn().mockResolvedValueOnce([{ name: 'Main', type: 'standard' }])
            .mockResolvedValueOnce(Array.from({ length: 1000 }, (_, i) => ({ name: `Other${i}`, type: 'standard' })));
        const explorer = create(list), [project, second] = await explorer.getChildren();
        const own = await explorer.getChildren(project), unaffected = await explorer.getChildren(second);
        const reads = vi.fn(() => other);
        for (const module of unaffected) { Object.defineProperty(module, 'filePath', { get: reads }); }
        const fired: unknown[] = [];
        const listener = explorer.onDidChangeTreeData(node => fired.push(node));
        explorer.refreshGitMarks(BOOK.toUpperCase().replaceAll('/', '\\'));
        listener.dispose();
        expect(fired).toEqual([...own, project]);
        expect(reads).not.toHaveBeenCalled();
        expect(list).toHaveBeenCalledTimes(2);
    });

    it('does not inspect another project while pruning removed modules', async () => {
        const other = 'C:/work/Other.vbp';
        host.findFiles.mockResolvedValue([{ scheme: 'file', fsPath: BOOK }, { scheme: 'file', fsPath: other }]);
        const list = vi.fn().mockResolvedValueOnce([{ name: 'Gone', type: 'standard' }])
            .mockResolvedValueOnce(Array.from({ length: 1000 }, (_, i) => ({ name: `Other${i}`, type: 'standard' })))
            .mockResolvedValue([]);
        const explorer = create(list), [project, second] = await explorer.getChildren();
        await explorer.getChildren(project);
        const unaffected = await explorer.getChildren(second), reads = vi.fn(() => other);
        for (const module of unaffected) { Object.defineProperty(module, 'filePath', { get: reads }); }
        explorer.retryLoad({ kind: 'loadError', label: 'Retry', filePath: BOOK });
        await explorer.getChildren(project);
        expect(explorer.getModuleNode(BOOK, 'Gone')).toBeUndefined();
        expect(explorer.getModuleNode(other, 'Other0')).toBe(unaffected[0]);
        expect(reads).not.toHaveBeenCalled();
        expect(list).toHaveBeenCalledTimes(3);
    });

    it('forgets removed modules and their procedure rows', async () => {
        const list = vi.fn().mockResolvedValueOnce([{ name: 'Gone', type: 'standard' }])
            .mockResolvedValue([{ name: 'Kept', type: 'standard' }]);
        const explorer = create(list), [project] = await explorer.getChildren();
        const [gone] = await explorer.getChildren(project);
        await explorer.getChildren(gone);
        explorer.retryLoad({ kind: 'loadError', label: 'Retry', filePath: BOOK });
        await explorer.getChildren(project);
        expect(explorer.getModuleNode(BOOK, 'Gone')).toBeUndefined();
        expect(explorer.getProcedureNode(BOOK, 'Gone', 'Sub Run')).toBeUndefined();
        expect(await explorer.resolveModuleNode(BOOK, 'Gone')).toBeUndefined();
        expect(await explorer.getChildren(gone)).toEqual([]);
        expect(explorer.getProcedureNode(BOOK, 'Gone', 'Sub Run')).toBeUndefined();
        expect(list).toHaveBeenCalledTimes(2);
    });

    it('updates module metadata while preserving the surviving row identity', async () => {
        const list = vi.fn().mockResolvedValueOnce([{
            name: 'Main', type: 'standard', documentType: 'worksheet', filePath: 'C:/work/old.bas',
        }]).mockResolvedValue([{
            name: 'MAIN', type: 'userform', filePath: 'C:/work/Main.frm',
        }]);
        const explorer = create(list), [project] = await explorer.getChildren();
        const [module] = await explorer.getChildren(project);
        await explorer.getChildren(module);
        explorer.retryLoad({ kind: 'loadError', label: 'Retry', filePath: BOOK });
        const [fresh] = await explorer.getChildren(project);
        expect(fresh).toBe(module);
        expect(fresh).toMatchObject({ label: 'MAIN', moduleName: 'MAIN', moduleType: 'userform', moduleFilePath: 'C:/work/Main.frm' });
        expect(fresh.documentType).toBeUndefined();
        expect((await explorer.getChildren(fresh)).map(row => row.kind)).toEqual(['designer', 'sub']);
    });

    it('drops the designer when a form is relisted as a standard module', async () => {
        const list = vi.fn().mockResolvedValueOnce([{ name: 'M', type: 'userform' }])
            .mockResolvedValue([{ name: 'M', type: 'standard' }]);
        const explorer = create(list), [project] = await explorer.getChildren();
        const [module] = await explorer.getChildren(project);
        expect((await explorer.getChildren(module))[0].kind).toBe('designer');
        explorer.retryLoad({ kind: 'loadError', label: 'Retry', filePath: BOOK });
        await explorer.getChildren(project);
        expect((await explorer.getChildren(module)).map(row => row.kind)).toEqual(['sub']);
    });

    for (const change of ['removed', 'form'] as const) {
        it(`does not restore stale procedure rows after a module is ${change}`, async () => {
            let release!: (subs: Array<{ name: string; kind: string; line: number }>) => void;
            const pending = new Promise<Array<{ name: string; kind: string; line: number }>>(resolve => { release = resolve; });
            const subs = vi.fn().mockReturnValueOnce(pending).mockResolvedValue([]);
            const list = vi.fn().mockResolvedValueOnce([{ name: 'M', type: 'standard' }])
                .mockResolvedValue(change === 'removed' ? [] : [{ name: 'M', type: 'userform' }]);
            const explorer = create(list, subs), [project] = await explorer.getChildren();
            const [module] = await explorer.getChildren(project);
            const loading = explorer.getChildren(module);
            explorer.retryLoad({ kind: 'loadError', label: 'Retry', filePath: BOOK });
            await explorer.getChildren(project);
            release([{ name: 'Old', kind: 'Sub', line: 1 }]);
            expect((await loading).map(row => row.kind)).toEqual(change === 'removed' ? [] : ['designer']);
            expect(explorer.getProcedureNode(BOOK, 'M', 'Sub Old')).toBeUndefined();
            expect(subs).toHaveBeenCalledTimes(change === 'removed' ? 1 : 2);
        });
    }
});
