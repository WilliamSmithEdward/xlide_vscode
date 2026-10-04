import { describe, expect, it, vi } from 'vitest';
vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());
import { ShapeRows } from '../src/shapeRows';
import type { ProjectEngine } from '../src/projectEngine';
import type { XlideNode } from '../src/projectExplorer';
const BOOK = '/work/Book.xlsm';
const project: XlideNode = { kind: 'project', label: 'Book', filePath: BOOK };
function fixture() {
    let name = 'Data', failed = false;
    const call = vi.fn(async (method: string) => {
        if (method === 'listWorkbookSheets') {
            if (failed) { throw new Error('Busy'); }
            return { sheets: [{ name, kind: 'worksheet' }] };
        }
        return { surfaces: [{ surface: name, shapes: [{ name: 'Box', kind: 'shape' }] }] };
    });
    const fire = vi.fn(), tree = new ShapeRows({ call } as unknown as ProjectEngine, fire);
    return { tree, fire, call, rename: (value: string) => { name = value; }, fail: (value: boolean) => { failed = value; } };
}

describe('retired worksheet rows', () => {
    it('keeps save notifications bounded after fifty sheet renames', async () => {
        const { tree, fire, rename } = fixture();
        const { folders: [sheets] } = await tree.projectRows(project, []);
        let sheet!: XlideNode, folder!: XlideNode;
        for (let i = 0; i < 50; i++) {
            rename(`Data ${i}`);
            tree.refresh(BOOK, { shapesChanged: true });
            [sheet] = await tree.children(sheets, async () => []);
            [folder] = await tree.children(sheet, async () => []);
            await tree.children(folder, async () => []);
        }
        fire.mockClear();
        for (let i = 0; i < 10; i++) { tree.refresh(BOOK, { shapesChanged: true }); }
        const notifications = fire.mock.calls.map(([node]) => node).filter(node => node !== sheets);
        expect(notifications).toHaveLength(20);
        expect(notifications).toEqual(
            Array.from({ length: 10 }, () => [sheet, folder]).flat());
    });

    it('does not revive retired rows when an old sheet name is reused', async () => {
        const { tree, fire, call, rename } = fixture();
        const { folders: [sheets] } = await tree.projectRows(project, []);
        const [oldSheet] = await tree.children(sheets, async () => []);
        const [oldFolder] = await tree.children(oldSheet, async () => []);
        await tree.children(oldFolder, async () => []);
        rename('Renamed');
        tree.refresh(BOOK);
        await tree.children(sheets, async () => []);
        rename('Data');
        tree.refresh(BOOK);
        const [fresh] = await tree.children(sheets, async () => []);
        expect(fresh).not.toBe(oldSheet);
        const before = call.mock.calls.length;
        expect(await tree.children(oldSheet, async () => [])).toEqual([]);
        expect(await tree.children(oldFolder, async () => [])).toEqual([]);
        expect(await tree.surfaceOf(oldFolder)).toBeUndefined();
        expect(call).toHaveBeenCalledTimes(before);
        fire.mockClear();
        tree.refresh(BOOK, { shapesChanged: true });
        expect(fire.mock.calls.flat()).not.toContain(oldFolder);
    });

    it('retains rows when a catalog fails and preserves case-only rename identity', async () => {
        const { tree, rename, fail } = fixture();
        const { folders: [sheets] } = await tree.projectRows(project, []);
        const [original] = await tree.children(sheets, async () => []);
        fail(true);
        tree.refresh(BOOK);
        await tree.catalog(BOOK);
        fail(false);
        rename('DATA');
        tree.refresh(BOOK);
        const [current] = await tree.children(sheets, async () => []);
        expect(current).toBe(original);
        expect(current.label).toBe('DATA');
    });

    it('does not recreate a retired folder from a shape read already in flight', async () => {
        const { tree, call, rename } = fixture();
        const { folders: [sheets] } = await tree.projectRows(project, []);
        const [old] = await tree.children(sheets, async () => []);
        let release!: (value: { surfaces: Array<{ surface: string; shapes: Array<{ name: string; kind: string }> }> }) => void;
        const listing = new Promise<{ surfaces: Array<{ surface: string; shapes: Array<{ name: string; kind: string }> }> }>(yes => { release = yes; });
        const original = call.getMockImplementation()!;
        call.mockImplementation(method => method === 'listShapes' ? listing : original(method));
        tree.refresh(BOOK, { shapesChanged: true });
        const pending = tree.children(old, async () => []);
        rename('Renamed');
        await tree.catalog(BOOK);
        release({ surfaces: [{ surface: 'Data', shapes: [{ name: 'Box', kind: 'shape' }] }] });
        expect(await pending).toEqual([]);
    });

    it('revokes group snapshots created just before their worksheet is retired', async () => {
        const { tree, call, rename } = fixture();
        const { folders: [sheets] } = await tree.projectRows(project, []);
        const [old] = await tree.children(sheets, async () => []);
        const original = call.getMockImplementation()!;
        call.mockImplementation(async method => method === 'listShapes'
            ? { surfaces: [{ surface: 'Data', shapes: [{ name: 'Pair', kind: 'group', shapes: [{ name: 'Part', kind: 'shape' }] }] }] }
            : original(method));
        tree.refresh(BOOK, { shapesChanged: true });
        const [folder] = await tree.children(old, async () => []);
        const [group] = await tree.children(folder, async () => []);
        const [member] = await tree.children(group, async () => []);
        rename('Renamed');
        await tree.catalog(BOOK);
        expect(await tree.children(group, async () => [])).toEqual([]);
        expect(tree.contextOf(group)).toBeUndefined();
        expect(tree.contextOf(member)).toBeUndefined();
    });
});
