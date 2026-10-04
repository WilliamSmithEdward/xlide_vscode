import { describe, expect, it, vi } from 'vitest';
vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());
import { ShapeRows } from '../src/shapeRows';
import type { ProjectEngine } from '../src/projectEngine';
import type { XlideNode } from '../src/projectExplorer';

describe('retained surface rows before their parent redraws', () => {
    it.each(['sheet', 'folder'] as const)('keeps shapes under a retained %s after a case-only rename', async target => {
        let name = 'Data';
        const call = vi.fn(async (method: string) => method === 'listWorkbookSheets'
            ? { sheets: [{ name, kind: 'worksheet' }] }
            : { surfaces: [{ surface: name, shapes: [{ name: 'Box', kind: 'shape' }] }] });
        const tree = new ShapeRows({ call } as unknown as ProjectEngine, () => {});
        const project: XlideNode = { kind: 'project', label: 'Book', filePath: '/work/Book.xlsm' };
        try {
            const { folders: [sheets] } = await tree.projectRows(project, []);
            const [sheet] = await tree.children(sheets, async () => []);
            const [folder] = await tree.children(sheet, async () => []);
            await tree.children(folder, async () => []);
            name = 'DATA';
            tree.refresh(project.filePath, { shapesChanged: true });
            const currentFolder = target === 'sheet' ? (await tree.children(sheet, async () => []))[0] : folder;
            expect(currentFolder).toBe(folder);
            const [box] = await tree.children(currentFolder, async () => []);
            expect(box.kind).toBe('shape');
            expect(box.label).toBe('Box');
            expect(tree.contextOf(box)).toMatchObject({ surface: 'DATA' });
            expect(currentFolder.surface).toBe('DATA');
            if (target === 'sheet') { expect(sheet.label).toBe('DATA'); }
            expect(call.mock.calls.filter(([method]) => method === 'listShapes')).toHaveLength(2);
            expect(call.mock.calls.filter(([method]) => method === 'listWorkbookSheets')).toHaveLength(1);
        } finally { tree.dispose(); }
    });

    it.each(['pptm', 'docm'])('keeps a retained %s surface current before its parent redraws', async extension => {
        let name = 'Data';
        const tree = new ShapeRows({ call: async () => ({ surfaces: [{ surface: name, shapes: [{ name: 'Box', kind: 'shape' }] }] }) } as unknown as ProjectEngine, () => {});
        const folder: XlideNode = { kind: 'shapes', shapeFolder: extension === 'pptm' ? 'slides' : 'module',
            moduleName: 'ThisDocument', label: 'Shapes', filePath: `/work/File.${extension}` };
        try {
            const [surface] = await tree.children(folder, async () => []);
            await tree.children(surface, async () => []);
            name = 'DATA';
            tree.refresh(folder.filePath);
            const [box] = await tree.children(surface, async () => []);
            expect(box.kind).toBe('shape');
            expect(box.label).toBe('Box');
            expect(surface.label).toBe('DATA');
            expect(tree.contextOf(surface)?.surface).toBe('DATA');
            expect(tree.contextOf(box)?.surface).toBe('DATA');
        } finally { tree.dispose(); }
    });
});
