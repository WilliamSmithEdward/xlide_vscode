import { describe, expect, it, vi } from 'vitest';
vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());
import { ShapeRows } from '../src/shapeRows';
import type { ProjectEngine } from '../src/projectEngine';
import type { XlideNode } from '../src/projectExplorer';
const BOOK = 'C:/work/Book.xlsm';
const project: XlideNode = { kind: 'project', label: 'Book', filePath: BOOK };
function moduleRow(): XlideNode {
    return { kind: 'module', label: 'Sheet1', moduleName: 'Sheet1', filePath: BOOK };
}
const catalog = { sheets: [{ name: 'Data', codeName: 'Sheet1', kind: 'worksheet' }] };
const shapes = (surface: string) => ({ surfaces: [{ surface, codeName: 'Sheet1', shapes: [{ name: 'Box', kind: 'shape' }] }] });
function create(call: ReturnType<typeof vi.fn>) {
    return new ShapeRows({ call } as unknown as ProjectEngine, () => {});
}
describe('sheet row context across refresh', () => {
    it.each(['renamed', 'removed'] as const)('resolves an already opened module folder after its sheet is %s', async change => {
        const call = vi.fn(async (method: string) => method === 'listWorkbookSheets' ? catalog : shapes('Data'));
        const tree = create(call), module = moduleRow();
        await tree.projectRows(project, [module]);
        const folder = await tree.moduleFolder(module);
        expect(folder).toBeDefined();
        await tree.children(folder!, async () => [module]);
        expect(await tree.surfaceOf(folder!)).toEqual({ host: 'excel', surface: 'Data' });
        call.mockResolvedValue(change === 'renamed' ? shapes('Renamed') : { surfaces: [] });
        tree.refresh(BOOK);
        const expected = change === 'renamed' ? { host: 'excel', surface: 'Renamed' } : undefined;
        expect(await tree.surfaceOf(folder!)).toEqual(expected);
        expect(await tree.surfaceOf(folder!)).toEqual(expected);
        expect(call.mock.calls.filter(([method]) => method === 'listShapes')).toHaveLength(2);
    });

    it.each(['success', 'failure'] as const)('does not return an overtaken add-shape lookup ending in %s', async outcome => {
        let resolve!: (value: unknown) => void, reject!: (error: Error) => void;
        const old = new Promise((resolveOld, rejectOld) => { resolve = resolveOld; reject = rejectOld; });
        const call = vi.fn().mockReturnValueOnce(old).mockResolvedValue(shapes('Renamed'));
        const tree = create(call);
        const folder: XlideNode = { kind: 'shapes', shapeFolder: 'module', moduleName: 'Sheet1', label: 'Shapes', filePath: BOOK };
        const pending = tree.surfaceOf(folder);
        tree.refresh(BOOK);
        expect(await tree.surfaceOf(folder)).toEqual({ host: 'excel', surface: 'Renamed' });
        if (outcome === 'success') { resolve(shapes('Old')); }
        else { reject(new Error('Obsolete failure')); }
        expect(await pending).toEqual({ host: 'excel', surface: 'Renamed' });
        expect(call).toHaveBeenCalledTimes(2);
    });

    it.each(['missing', 'failed'] as const)('detaches modules from obsolete sheet parents when the catalog is %s', async change => {
        const call = vi.fn().mockResolvedValue(catalog), tree = create(call), module = moduleRow();
        const { folders: [folder] } = await tree.projectRows(project, [module]);
        expect(tree.parentOf(module)).toBe(folder);
        if (change === 'failed') { call.mockRejectedValue(new Error('Workbook busy')); }
        else { call.mockResolvedValue({ sheets: [] }); }
        tree.refresh(BOOK);
        const rows = await tree.projectRows(project, [module]);
        expect(rows.modules).toEqual([module]);
        expect(tree.parentOf(module)).toBeUndefined();
        expect(module.label).toBe('Sheet1');
        expect(module.sheetName).toBeUndefined();
        call.mockResolvedValue(catalog);
        tree.refresh(BOOK);
        await tree.projectRows(project, [module]);
        expect(tree.parentOf(module)).toBe(folder);
        expect(module.label).toBe('Sheet1 (Data)');
    });
});
