import { describe, expect, it, vi } from 'vitest';
vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());
import { ShapeRows } from '../src/shapeRows';
import type { ProjectEngine } from '../src/projectEngine';
import type { XlideNode } from '../src/projectExplorer';
import type { ShapeSurface } from '../src/vba/projectService';
import type { WorkbookSheet } from '../src/vba/workbookSheets';

const project: XlideNode = { kind: 'project', label: 'Book', filePath: 'C:\\work\\Book.xlsm' };
function setup(sheets: WorkbookSheet[], surfaces: ShapeSurface[]) {
    const call = vi.fn(async (method: string) => {
        if (method === 'listWorkbookSheets') { return { sheets }; }
        if (method === 'listShapes') { return { surfaces }; }
        throw new Error(method);
    });
    return { tree: new ShapeRows({ call } as unknown as ProjectEngine, () => {}), call };
}
const shape = { name: 'Box', kind: 'shape' as const };

describe('Sheets folder surface lookup', () => {
    it('reads each surface name once per expansion, including cached expansions', async () => {
        const sheets = Array.from({ length: 1000 }, (_, i) => ({ name: 'Sheet' + i, kind: 'worksheet' as const }));
        let reads = 0;
        const surfaces = sheets.map(sheet => ({ get surface() { reads++; return sheet.name; }, shapes: [shape] }));
        const { tree, call } = setup(sheets, surfaces);
        const { folders: [folder] } = await tree.projectRows(project, []);
        for (let expansion = 0; expansion < 2; expansion++) {
            reads = 0;
            const rows = await tree.children(folder, async () => []);
            expect(reads).toBeLessThanOrEqual(surfaces.length);
            expect(rows.map(row => row.label)).toEqual(sheets.map(sheet => sheet.name));
            expect(rows.every(row => tree.parentOf(row) === folder && row.itemCount === 1)).toBe(true);
        }
        expect(call.mock.calls.map(args => args[0])).toEqual(['listWorkbookSheets', 'listShapes']);
    });

    it('keeps the first exact matching surface, missing sheets, and module parents', async () => {
        const sheets: WorkbookSheet[] = [
            { name: 'First', kind: 'worksheet', codeName: 'M1' },
            { name: 'Second', kind: 'worksheet', codeName: 'M2' },
            { name: 'Third', kind: 'worksheet' },
            { name: 'Fourth', kind: 'chartsheet', state: 'hidden' },
        ];
        const surfaces = Object.freeze([
            Object.freeze({ surface: 'First', shapes: [] }),
            Object.freeze({ surface: 'First', shapes: [shape] }),
            Object.freeze({ surface: 'second', shapes: [shape] }),
            Object.freeze({ surface: 'Third', shapes: [shape] }),
            Object.freeze({ surface: 'Unlisted', shapes: [shape] }),
        ]) as unknown as ShapeSurface[];
        const modules: XlideNode[] = ['M1', 'M2'].map(moduleName => ({ kind: 'module', label: moduleName, moduleName, hasCode: false, filePath: project.filePath }));
        const { tree } = setup(sheets, surfaces);
        const { folders: [folder] } = await tree.projectRows(project, modules);
        const [third, bare] = await tree.children(folder, async () => modules);
        expect(third.label).toBe('Third');
        expect(third.itemCount).toBe(1);
        expect(tree.parentOf(third)).toBe(folder);
        const bareRows = await tree.children(bare, async () => modules);
        expect(bareRows.map(row => row.label)).toEqual(['M1 (First)', 'M2 (Second)', 'Fourth']);
        expect(bareRows.slice(0, 2)).toEqual(modules);
        expect(bareRows.every(row => tree.parentOf(row) === bare)).toBe(true);
        expect(tree.treeItem(bareRows[2])?.description).toContain('hidden');
    });

    it('uses refreshed shape contents to move an empty module into the bare folder', async () => {
        const sheets: WorkbookSheet[] = [{ name: 'Data', codeName: 'Sheet1', kind: 'worksheet' }];
        const surfaces: ShapeSurface[] = [{ surface: 'Data', shapes: [shape] }];
        const modules: XlideNode[] = [{ kind: 'module', label: 'Sheet1', moduleName: 'Sheet1', hasCode: false, filePath: project.filePath }];
        const { tree, call } = setup(sheets, surfaces);
        const { folders: [folder] } = await tree.projectRows(project, modules);
        expect(await tree.children(folder, async () => modules)).toEqual(modules);
        call.mockImplementation(async method => method === 'listWorkbookSheets' ? { sheets } : { surfaces: [{ surface: 'Data', shapes: [] }] });
        tree.refresh(project.filePath, { shapesChanged: true });
        const [bare] = await tree.children(folder, async () => modules);
        expect(bare.shapeFolder).toBe('bareSheets');
        expect(await tree.children(bare, async () => modules)).toEqual(modules);
        expect(tree.parentOf(modules[0])).toBe(bare);
    });
});
