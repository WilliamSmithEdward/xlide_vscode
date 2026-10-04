import { describe, expect, it, vi } from 'vitest';
vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());
import { ShapeRows } from '../src/shapeRows';
import type { ProjectEngine } from '../src/projectEngine';
import type { XlideNode } from '../src/projectExplorer';

const project: XlideNode = { kind: 'project', label: 'Book', filePath: 'C:/work/Book.xlsm' };

describe('shape tree item identity', () => {
    it.each([true, false])('preserves a sheet ID across case-only renames with shapes=%s', async hasShapes => {
        let name = 'Data';
        const call = vi.fn(async (method: string) => method === 'listWorkbookSheets'
            ? { sheets: [{ name, kind: 'worksheet' }] }
            : { surfaces: [{ surface: name, shapes: hasShapes ? [{ name: 'Box', kind: 'shape' }] : [] }] });
        const tree = new ShapeRows({ call } as unknown as ProjectEngine, () => {});
        try {
            const { folders: [sheets] } = await tree.projectRows(project, []);
            const sheetRow = async () => {
                const [row] = await tree.children(sheets, async () => []);
                return hasShapes ? row : (await tree.children(row, async () => []))[0];
            };
            const original = await sheetRow(), id = tree.treeItem(original).id;
            name = 'DATA';
            tree.refresh(project.filePath);
            const renamed = await sheetRow();
            expect(renamed).toBe(original);
            expect(renamed.label).toBe('DATA');
            expect(tree.treeItem(renamed).id).toBe(id);
            expect(tree.contextOf(renamed)?.surface).toBe('DATA');
        } finally { tree.dispose(); }
    });

    it('preserves slide and shape IDs when only their letter case changes', async () => {
        let surface = 'Slide 1', name = 'Box';
        const tree = new ShapeRows({ call: async () => ({ surfaces: [{ surface, shapes: [{ name, kind: 'shape' }] }] }) } as unknown as ProjectEngine, () => {});
        const slides: XlideNode = { kind: 'shapes', shapeFolder: 'slides', label: 'Slides', filePath: 'C:/work/Deck.pptm' };
        try {
            const [slide] = await tree.children(slides, async () => []);
            const [shape] = await tree.children(slide, async () => []);
            const ids = [slide, shape].map(row => tree.treeItem(row).id);
            surface = 'SLIDE 1'; name = 'BOX';
            tree.refresh(slides.filePath);
            const [renamedSlide] = await tree.children(slides, async () => []);
            const [renamedShape] = await tree.children(renamedSlide, async () => []);
            expect([renamedSlide, renamedShape].map(row => tree.treeItem(row).id)).toEqual(ids);
            expect(tree.contextOf(renamedShape)).toMatchObject({ surface: 'SLIDE 1', shape: { name: 'BOX' } });
        } finally { tree.dispose(); }
    });

    it('distinguishes a slash in a shape name from a group path separator', async () => {
        const tree = new ShapeRows({ call: async () => ({ surfaces: [{ surface: 'Data', shapes: [
            { name: 'A/B', kind: 'group', shapes: [{ name: 'Top member', kind: 'shape' }] },
            { name: 'A', kind: 'group', shapes: [{ name: 'B', kind: 'group', shapes: [{ name: 'Nested member', kind: 'shape' }] }] },
        ] }] }) } as unknown as ProjectEngine, () => {});
        const folder: XlideNode = { kind: 'shapes', shapeFolder: 'surface', surface: 'Data', label: 'Shapes', filePath: project.filePath };
        try {
            const [top, group] = await tree.children(folder, async () => []);
            const [nested] = await tree.children(group, async () => []);
            expect(top.shapePath).toEqual(['A/B']);
            expect(nested.shapePath).toEqual(['A', 'B']);
            expect(tree.treeItem(top).id).not.toBe(tree.treeItem(nested).id);
            expect((await tree.children(top, async () => []))[0].label).toBe('Top member');
            expect((await tree.children(nested, async () => []))[0].label).toBe('Nested member');
        } finally { tree.dispose(); }
    });
});
