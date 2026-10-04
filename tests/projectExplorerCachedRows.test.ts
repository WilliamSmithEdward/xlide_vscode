import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
const host = vi.hoisted(() => ({ findFiles: vi.fn() }));
vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
    workspace: { findFiles: host.findFiles, workspaceFolders: [{ uri: { fsPath: 'C:/work' } }] },
}));
import { ProjectExplorer } from '../src/projectExplorer';
const BOOK = 'C:/work/App.vbp';
let explorers: ProjectExplorer[] = [];
function create(list: () => Promise<unknown>) {
    const call = vi.fn((method: string) => method === 'listModules' ? list() : Promise.resolve({ isPasswordProtected: false, isSigned: false }));
    const explorer = new ProjectExplorer({ call } as unknown as ConstructorParameters<typeof ProjectExplorer>[0]);
    explorers.push(explorer);
    return explorer;
}
function deferred<T>() {
    let resolve!: (value: T) => void;
    const promise = new Promise<T>(yes => { resolve = yes; });
    return { promise, resolve };
}
beforeEach(() => {
    explorers = [];
    host.findFiles.mockResolvedValue([{ scheme: 'file', fsPath: BOOK }]);
});
afterEach(() => { for (const explorer of explorers) explorer.dispose(); });

function fixture() {
    let reads = 0;
    const modules = Object.freeze(Array.from({ length: 1000 }, (_, i) => Object.freeze({
        get name() { reads++; return 'M' + String((i * 317) % 1000).padStart(4, '0'); },
        type: ['standard', 'class', 'userform'][i % 3],
    })));
    return { modules, reset: () => { reads = 0; }, reads: () => reads };
}

describe('cached module rows', () => {
    it('does not revisit backend metadata during 200 cached project expansions', async () => {
        const data = fixture(), list = vi.fn(async () => data.modules), explorer = create(list);
        const [project] = await explorer.getChildren(), original = await explorer.getChildren(project);
        const ids = original.map(node => explorer.getTreeItem(node).id);
        data.reset();
        for (let i = 0; i < 200; i++) {
            const rows = await explorer.getChildren(project);
            expect(rows).toEqual(original);
            expect(rows[0]).toBe(original[0]);
        }
        expect(original.map(node => explorer.getTreeItem(node).id)).toEqual(ids);
        expect(data.reads()).toBe(0);
        expect(list).toHaveBeenCalledTimes(1);
    });

    it.each([2, 8])('shares row construction for %i overlapping expansion/follow callers', async callers => {
        async function measure(count: number) {
            const data = fixture(), pending = deferred<typeof data.modules>();
            const list = vi.fn(() => pending.promise), explorer = create(list);
            const [project] = await explorer.getChildren();
            const expansion = explorer.getChildren(project);
            const follow = Array.from({ length: count - 1 }, () => explorer.resolveModuleNode(BOOK, 'M0317'));
            await Promise.resolve(); await Promise.resolve(); pending.resolve(data.modules);
            const rows = await expansion;
            for (const node of await Promise.all(follow)) expect(node).toBe(rows.find(row => row.moduleName === 'M0317'));
            expect(list).toHaveBeenCalledTimes(1);
            return data.reads();
        }
        expect(await measure(callers)).toBe(await measure(1));
    });

    it('updates cached rows for editor folder edits and re-reads after the editor closes', async () => {
        const list = vi.fn().mockResolvedValueOnce([{ name: 'M', type: 'standard', folder: 'Saved' }])
            .mockResolvedValue([{ name: 'M', type: 'standard', folder: 'AfterSave' }]);
        const explorer = create(list); explorer.setView('folders');
        const [project] = await explorer.getChildren(); await explorer.getChildren(project);
        const module = explorer.getModuleNode(BOOK, 'M')!;
        explorer.setModuleFolder(BOOK, 'M', 'Edited');
        expect((await explorer.getChildren(project)).map(node => node.folder)).toEqual(['Edited']);
        expect(explorer.getModuleNode(BOOK, 'M')).toBe(module); expect(module.folder).toBe('Edited');
        expect(list).toHaveBeenCalledTimes(1);
        explorer.forgetModuleFolder(BOOK, 'M');
        expect((await explorer.getChildren(project)).map(node => node.folder)).toEqual(['AfterSave']);
        expect(module.folder).toBe('AfterSave'); expect(list).toHaveBeenCalledTimes(2);
        explorer.refresh(); const [fresh] = await explorer.getChildren(); await explorer.getChildren(fresh);
        expect(explorer.getModuleNode(BOOK, 'M')).not.toBe(module); expect(list).toHaveBeenCalledTimes(3);
    });

    it('does not let an older load mutate current cached rows for the same module', async () => {
        const old = deferred<Array<{ name: string; type: string; folder: string; hasCode: boolean }>>();
        const list = vi.fn().mockReturnValueOnce(old.promise).mockResolvedValue([{ name: 'M', type: 'standard', folder: 'New', hasCode: true }]);
        const explorer = create(list), [project] = await explorer.getChildren();
        const pending = explorer.getChildren(project); explorer.refresh();
        const [current] = await explorer.getChildren(), rows = await explorer.getChildren(current);
        old.resolve([{ name: 'M', type: 'standard', folder: 'Old', hasCode: false }]);
        const result = await pending;
        expect(result[0]).toBe(rows[0]); expect(rows[0]).toMatchObject({ folder: 'New', hasCode: true });
        expect((await explorer.getChildren(current))[0]).toBe(rows[0]); expect(list).toHaveBeenCalledTimes(2);
    });
});
