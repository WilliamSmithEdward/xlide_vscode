import { describe, expect, it, vi } from 'vitest';
vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());
import { ShapeRows } from '../src/shapeRows';
import type { ProjectEngine } from '../src/projectEngine';
import type { XlideNode } from '../src/projectExplorer';

const BOOK = 'C:/work/Book.xlsm';
const project: XlideNode = { kind: 'project', label: 'Book', filePath: BOOK };
async function fixture() {
    let name = 'Data', hasShapes = false, removed = false;
    const call = vi.fn(async (method: string) => method === 'listWorkbookSheets'
        ? { sheets: removed ? [] : [{ name, kind: 'worksheet' }] }
        : { surfaces: hasShapes ? [{ surface: name, shapes: [{ name: 'Box', kind: 'shape' }] }] : [] });
    const tree = new ShapeRows({ call } as unknown as ProjectEngine, () => {});
    const { folders: [sheets] } = await tree.projectRows(project, []);
    const [bare] = await tree.children(sheets, async () => []);
    return { tree, call, sheets, bare, rename: () => { name = 'Renamed'; },
        addShape: () => { hasShapes = true; }, remove: () => { removed = true; } };
}

describe('retained bare-sheet folders after refresh', () => {
    it('reads the renamed sheet when the bare folder is requested before its parent redraws', async () => {
        const f = await fixture();
        f.rename(); f.tree.refresh(BOOK);
        expect((await f.tree.children(f.bare, async () => [])).map(row => row.label)).toEqual(['Renamed']);
    });

    it('drops a sheet that gained its first shape before the parent redraws', async () => {
        const f = await fixture();
        f.addShape(); f.tree.refresh(BOOK, { shapesChanged: true });
        expect(await f.tree.children(f.bare, async () => [])).toEqual([]);
    });

    it('drops the final bare row after Sheets has redrawn without the folder', async () => {
        const f = await fixture();
        f.addShape(); f.tree.refresh(BOOK, { shapesChanged: true });
        await f.tree.children(f.sheets, async () => []);
        expect(await f.tree.children(f.bare, async () => [])).toEqual([]);
    });

    it('does not return a deleted sheet from a retained bare folder', async () => {
        const f = await fixture();
        f.remove(); f.tree.refresh(BOOK);
        expect(await f.tree.children(f.bare, async () => [])).toEqual([]);
    });

    it('does not resurrect old rows after a full clear', async () => {
        const f = await fixture();
        f.tree.clear();
        const before = f.call.mock.calls.length;
        expect(await f.tree.children(f.bare, async () => [])).toEqual([]);
        expect(f.call.mock.calls.length).toBe(before);
    });

    it('keeps repeated bare-folder expansion cached', async () => {
        const f = await fixture();
        const rows = await f.tree.children(f.bare, async () => []);
        const modules = vi.fn(async () => []);
        const before = f.call.mock.calls.length;
        for (let i = 0; i < 200; i++) { expect(await f.tree.children(f.bare, modules)).toBe(rows); }
        expect(f.call.mock.calls.length).toBe(before);
        expect(modules).not.toHaveBeenCalled();
    });

    it('joins a newer refresh while rebuilding the bare folder', async () => {
        const f = await fixture();
        let release!: (rows: XlideNode[]) => void;
        const modules = vi.fn().mockReturnValueOnce(new Promise<XlideNode[]>(yes => { release = yes; }))
            .mockResolvedValue([]);
        f.tree.refresh(BOOK);
        const pending = f.tree.children(f.bare, modules);
        expect(modules).toHaveBeenCalledTimes(1);
        f.rename(); f.tree.refresh(BOOK);
        release([]);
        expect((await pending).map(row => row.label)).toEqual(['Renamed']);
        expect(modules).toHaveBeenCalledTimes(2);
    });

    it('returns no retained rows if disposed during reclassification', async () => {
        const f = await fixture();
        let release!: (rows: XlideNode[]) => void;
        f.tree.refresh(BOOK);
        const before = f.call.mock.calls.length;
        const pending = f.tree.children(f.bare, () => new Promise<XlideNode[]>(yes => { release = yes; }));
        f.tree.dispose();
        release([]);
        expect(await pending).toEqual([]);
        expect(f.call.mock.calls.length).toBe(before);
    });
});
