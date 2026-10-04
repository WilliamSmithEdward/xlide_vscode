import { describe, expect, it, vi } from 'vitest';
vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());
import { ShapeRows } from '../src/shapeRows';
import type { ProjectEngine } from '../src/projectEngine';
import type { XlideNode } from '../src/projectExplorer';
const BOOK = 'C:/work/Book.xlsm';
const project: XlideNode = { kind: 'project', label: 'Book', filePath: BOOK };
const moduleRow = (): XlideNode => ({ kind: 'module', label: 'Sheet1', moduleName: 'Sheet1', sheetName: 'Data', filePath: BOOK });
const sheets = (name: string) => ({ sheets: [{ name, codeName: 'Sheet1', kind: 'worksheet' }] });
const shapes = (name: string) => ({ surfaces: [{ surface: 'Data', codeName: 'Sheet1', shapes: [{ name, kind: 'shape' }] }] });
function deferred() {
    let resolve!: (value: unknown) => void, reject!: (error: Error) => void;
    const promise = new Promise<unknown>((yes, no) => { resolve = yes; reject = no; });
    return { promise, resolve, reject };
}
function create(call: ReturnType<typeof vi.fn>) {
    return new ShapeRows({ call } as unknown as ProjectEngine, () => {});
}
describe('shape row rendering lifetime', () => {
    it('does not let an old catalog render overwrite a newly renamed module', async () => {
        const old = deferred(), module = moduleRow();
        const tree = create(vi.fn().mockReturnValueOnce(old.promise).mockResolvedValue(sheets('New')));
        const pending = tree.projectRows(project, [module]);
        tree.refresh(BOOK);
        const fresh = await tree.projectRows(project, [module]);
        old.resolve(sheets('Old'));
        expect(await pending).toEqual(fresh);
        expect(module.label).toBe('Sheet1 (New)');
        expect(module.sheetName).toBe('New');
    });
    it.each(['success', 'failure'] as const)('joins current shape rows after an obsolete read ends in %s', async outcome => {
        const old = deferred();
        const call = vi.fn().mockReturnValueOnce(old.promise).mockResolvedValue(shapes('New'));
        const tree = create(call);
        const folder: XlideNode = { kind: 'shapes', shapeFolder: 'surface', surface: 'Data', label: 'Shapes', filePath: BOOK };
        const pending = tree.children(folder, async () => []);
        tree.refresh(BOOK);
        expect((await tree.children(folder, async () => [])).map(row => row.label)).toEqual(['New']);
        if (outcome === 'success') { old.resolve(shapes('Old')); }
        else { old.reject(new Error('obsolete failure')); }
        expect((await pending).map(row => row.label)).toEqual(['New']);
        expect(call).toHaveBeenCalledTimes(2);
    });
    it('keeps surface rows accurate after a rename that only changes case', async () => {
        const call = vi.fn(async (method: string) => method === 'listWorkbookSheets'
            ? { sheets: [{ name: 'Data', kind: 'worksheet' }] } : shapes('Box'));
        const tree = create(call);
        const { folders: [folder] } = await tree.projectRows(project, []);
        const [before] = await tree.children(folder, async () => []);
        call.mockImplementation(async (method: string) => method === 'listWorkbookSheets'
            ? { sheets: [{ name: 'DATA', kind: 'worksheet' }] }
            : { surfaces: [{ surface: 'DATA', shapes: [{ name: 'Box', kind: 'shape' }] }] });
        tree.refresh(BOOK);
        const [after] = await tree.children(folder, async () => []);
        expect(after).toBe(before);
        expect(after.label).toBe('DATA');
        expect(after.surface).toBe('DATA');
        const [shapeFolder] = await tree.children(after, async () => []);
        expect(shapeFolder?.shapeFolder).toBe('surface');
        expect((await tree.children(shapeFolder, async () => [])).map(row => row.label)).toEqual(['Box']);
    });

    it('rechecks the sheet catalog when a refresh lands during its shape read', async () => {
        const old = deferred(), module = moduleRow();
        let renamed = false, shapeCalls = 0;
        const call = vi.fn(async (method: string) => method === 'listWorkbookSheets'
            ? sheets(renamed ? 'New' : 'Old') : ++shapeCalls === 1 ? old.promise : { surfaces: [] });
        const tree = create(call);
        const { folders: [folder] } = await tree.projectRows(project, [module]);
        const pending = tree.children(folder, async () => [module]);
        for (let i = 0; i < 10 && shapeCalls === 0; i++) { await Promise.resolve(); }
        expect(shapeCalls).toBe(1);
        renamed = true;
        tree.refresh(BOOK);
        await tree.children(folder, async () => [module]);
        old.resolve({ surfaces: [] });
        expect(await pending).toEqual([module]);
        expect(module.label).toBe('Sheet1 (New)');
        expect(call.mock.calls.filter(([method]) => method === 'listWorkbookSheets')).toHaveLength(2);
        expect(shapeCalls).toBe(2);
    });

    it('does not reread an unrelated project when another workbook refreshes', async () => {
        const old = deferred(), module = moduleRow();
        const call = vi.fn().mockReturnValue(old.promise), tree = create(call);
        const pending = tree.projectRows(project, [module]);
        tree.refresh('C:/work/Other.xlsm');
        old.resolve(sheets('Data'));
        await pending;
        expect(call).toHaveBeenCalledTimes(1);
        expect(module.label).toBe('Sheet1 (Data)');
    });
});
