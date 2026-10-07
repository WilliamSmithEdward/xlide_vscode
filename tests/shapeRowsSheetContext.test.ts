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
    it.each(['case', 'renamed', 'removed', 'chart', 'failed'] as const)(
        'rechecks module-less sheet and folder targets after %s changes', async change => {
            let name = 'Data', kind = 'worksheet', removed = false, failed = false;
            const call = vi.fn(async (method: string) => {
                if (method === 'listWorkbookSheets') {
                    if (failed) { throw new Error('Workbook busy'); }
                    return { sheets: removed ? [] : [{ name, kind }] };
                }
                return shapes(name);
            });
            const tree = create(call);
            const { folders: [sheets] } = await tree.projectRows(project, []);
            const [sheet] = await tree.children(sheets, async () => []);
            const [folder] = await tree.children(sheet, async () => []);
            name = change === 'case' ? 'DATA' : change === 'renamed' ? 'Renamed' : 'Data';
            kind = change === 'chart' ? 'chartsheet' : 'worksheet';
            removed = change === 'removed';
            failed = change === 'failed';
            tree.refresh(BOOK);
            const expected = change === 'case' ? { host: 'excel', surface: 'DATA' } : undefined;
            expect(await tree.surfaceOf(sheet)).toEqual(expected);
            expect(await tree.surfaceOf(folder)).toEqual(expected);
            expect(call.mock.calls.filter(([method]) => method === 'listShapes')).toHaveLength(1);
        });

    it('adds to an empty module-less worksheet using the catalog without reading drawings', async () => {
        const call = vi.fn(async (method: string) => method === 'listWorkbookSheets'
            ? { sheets: [{ name: 'Data', kind: 'worksheet' }] } : { surfaces: [] });
        const tree = create(call);
        const { folders: [sheets] } = await tree.projectRows(project, []);
        const [bare] = await tree.children(sheets, async () => []);
        const [sheet] = await tree.children(bare, async () => []);
        tree.refresh(BOOK);
        expect(await tree.surfaceOf(sheet)).toEqual({ host: 'excel', surface: 'Data' });
        expect(call.mock.calls.filter(([method]) => method === 'listShapes')).toHaveLength(1);
    });

    it('follows a retained sheet code name instead of a replacement with the old display name', async () => {
        const call = vi.fn(async (method: string) => method === 'listWorkbookSheets' ? catalog : shapes('Data'));
        const tree = create(call);
        const { folders: [sheets] } = await tree.projectRows(project, []);
        const [sheet] = await tree.children(sheets, async () => []);
        const [folder] = await tree.children(sheet, async () => []);
        call.mockResolvedValue({ sheets: [
            { name: 'Data', codeName: 'Replacement', kind: 'worksheet' },
            { name: 'Renamed', codeName: 'Sheet1', kind: 'worksheet' },
        ] });
        tree.refresh(BOOK);
        for (const node of [sheet, folder]) {
            expect(await tree.surfaceOf(node)).toEqual({ host: 'excel', surface: 'Renamed' });
        }
    });

    it.each(['success', 'failure'] as const)('retries an overtaken worksheet catalog ending in %s', async outcome => {
        let resolve!: (value: unknown) => void, reject!: (error: Error) => void;
        const old = new Promise((yes, no) => { resolve = yes; reject = no; });
        const call = vi.fn().mockReturnValueOnce(old).mockResolvedValue({ sheets: [{ name: 'DATA', kind: 'worksheet' }] });
        const tree = create(call);
        const sheet: XlideNode = { kind: 'surface', label: 'Data', surface: 'Data', filePath: BOOK };
        const pending = tree.surfaceOf(sheet);
        tree.refresh(BOOK);
        expect(await tree.surfaceOf(sheet)).toEqual({ host: 'excel', surface: 'DATA' });
        if (outcome === 'success') { resolve({ sheets: [{ name: 'Data', kind: 'worksheet' }] }); }
        else { reject(new Error('Obsolete catalog')); }
        expect(await pending).toEqual({ host: 'excel', surface: 'DATA' });
        expect(call).toHaveBeenCalledTimes(2);
    });

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
