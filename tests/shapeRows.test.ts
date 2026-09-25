import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

const vscodeMock = vi.hoisted(() => ({
    findFiles: vi.fn(),
    fired: [] as unknown[],
}));

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
    EventEmitter: class {
        event = vi.fn();
        fire = vi.fn((node?: unknown) => {
            vscodeMock.fired.push(node);
        });
        dispose = vi.fn();
    },
    workspace: {
        findFiles: vscodeMock.findFiles,
        workspaceFolders: [{ uri: { fsPath: 'C:\\work' } }],
    },
}));

import { ProjectExplorer, type XlideNode } from '../src/projectExplorer';
import { ProjectEngine } from '../src/projectEngine';
import { editShape, listShapes } from '../src/vba/projectService';

// The sheet and shape rows, drawn from files the applications saved and read
// by the real engine: where each host's sheets and shapes hang, what each
// row offers, and that a changed file shows through. SheetsFixture is one
// workbook Excel saved three ways: Budget, whose module is Sheet1; Drawn,
// with a shape and no module; Trend, a chart sheet; Later; and Hidden.
const FIXTURES = path.join(__dirname, 'fixtures', 'binaries');

let dir: string;
let engine: ProjectEngine;
let explorer: ProjectExplorer;

beforeEach(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-shape-rows-'));
    vscodeMock.fired = [];
    engine = new ProjectEngine({} as never);
    explorer = new ProjectExplorer(engine);
});

afterEach(() => {
    fs.rmSync(dir, { recursive: true, force: true });
});

/** The project row of a copy of a fixture. */
async function projectOf(fixture: string, as = fixture): Promise<XlideNode> {
    const file = path.join(dir, as);
    fs.copyFileSync(path.join(FIXTURES, fixture), file);
    vscodeMock.findFiles.mockResolvedValue([{ scheme: 'file', fsPath: file }]);
    const [project] = await explorer.getChildren();
    return project;
}

/** Waits for the rows' background reads to land. */
const settle = (): Promise<void> => new Promise((resolve) => setTimeout(resolve, 50));

const say = (node: XlideNode): string => {
    const item = explorer.getTreeItem(node);
    return [node.label, item.description, item.contextValue].filter(Boolean).join(' | ');
};

/** The Sheets folder of a workbook, and the rows under it. */
async function sheetsOf(project: XlideNode): Promise<{ sheets: XlideNode; rows: XlideNode[] }> {
    const [sheets] = await explorer.getChildren(project);
    expect(say(sheets)).toMatch(/^Sheets \| \d+ sheets? \| sheetsFolder$/);
    return { sheets, rows: await explorer.getChildren(sheets) };
}

describe('a workbook\'s sheets in the explorer', () => {
    it('lead the project as a Sheets folder, and its modules stay where they were', async () => {
        const project = await projectOf('SheetsFixture.xlsm');
        const rows = await explorer.getChildren(project);

        expect(say(rows[0])).toBe('Sheets | 5 sheets | sheetsFolder');
        const modules = rows.slice(1).map((row) => row.moduleName);
        expect(modules).toContain('ThisWorkbook');
        expect(modules).not.toContain('Sheet1');
        expect(rows.slice(1).every((row) => row.kind === 'module')).toBe(true);
        expect(explorer.getParent(rows[0])).toBe(project);
    });

    it('list every sheet in tab order: the module for one that has it, named for its sheet, the sheet itself for one with shapes, and the rest in a folder at the end', async () => {
        const project = await projectOf('SheetsFixture.xlsm');
        const { sheets, rows } = await sheetsOf(project);

        expect(rows.map(say)).toEqual([
            'Sheet1 (Budget) | document | module-document',
            'Drawn | shapeSurface-add',
            'Sheets With No Modules or Shapes | 3 sheets | shapeFolder',
        ]);
        const [budget, drawn, bare] = rows;
        expect(budget).toMatchObject({ kind: 'module', moduleName: 'Sheet1', sheetName: 'Budget' });
        expect(explorer.getParent(budget)).toBe(sheets);
        expect(explorer.getParent(drawn)).toBe(sheets);
        expect(explorer.getParent(bare)).toBe(sheets);

        const rest = await explorer.getChildren(bare);
        expect(rest.map(say)).toEqual(['Trend | chart sheet | shapeSurface', 'Later | shapeSurface-add', 'Hidden | hidden | shapeSurface-add']);
        expect(rest.map((row) => explorer.getTreeItem(row).collapsibleState)).toEqual([0, 0, 0]);
        expect(explorer.getParent(rest[1])).toBe(bare);
        expect(explorer.getTreeItem(rest[0]).tooltip).toBe('Trend: a chart sheet with no module in the VBA project.');
    });

    it('draw a Shapes folder under a sheet only when it has shapes', async () => {
        const project = await projectOf('SheetsFixture.xlsm');
        const { rows: [budget, drawn] } = await sheetsOf(project);

        // Budget has a module and no shapes: its procedures, and no folder.
        expect((await explorer.getChildren(budget)).some((row) => row.kind === 'shapes')).toBe(false);

        expect(explorer.getTreeItem(drawn).collapsibleState).toBe(1);
        expect(explorer.getTreeItem(drawn).tooltip).toBe('Drawn: a worksheet with no module in the VBA project, with 1 shape.');
        const [folder] = await explorer.getChildren(drawn);
        expect(say(folder)).toBe('Shapes | shapeFolder-add');
        expect(explorer.getParent(folder)).toBe(drawn);
        expect((await explorer.getChildren(folder)).map(say)).toEqual(['DrawnBox | AutoShape | shape-edit-link-delete']);
        expect(await explorer.shapeSurfaceOf(folder)).toEqual({ host: 'excel', surface: 'Drawn' });
    });

    it.each(['xlsb', 'xls'])('are listed the same from a .%s, whose shapes the tree cannot read, so every sheet without a module sits in the folder', async (extension) => {
        const project = await projectOf(`SheetsFixture.${extension}`);
        const { rows } = await sheetsOf(project);

        expect(rows.map(say)).toEqual([
            'Sheet1 (Budget) | document | module-document',
            'Sheets With No Modules or Shapes | 4 sheets | shapeFolder',
        ]);
        const rest = await explorer.getChildren(rows[1]);
        // No shape tools for the format: nothing to add from a row.
        expect(rest.map(say)).toEqual(['Drawn | shapeSurface', 'Trend | chart sheet | shapeSurface', 'Later | shapeSurface', 'Hidden | hidden | shapeSurface']);
        expect((await explorer.getChildren(rows[0])).some((row) => row.kind === 'shapes')).toBe(false);
    });

    it('hang a worksheet\'s shapes under its module, above its procedures', async () => {
        const project = await projectOf('ShapesFixture.xlsm');
        const { sheets, rows } = await sheetsOf(project);
        const sheet1 = rows.find((row) => row.moduleName === 'Sheet1')!;
        expect(say(sheet1)).toBe('Sheet1 (Sheet1) | document | module-document');
        const [folder] = await explorer.getChildren(sheet1);
        expect(say(folder)).toBe('Shapes | shapeFolder-add');

        const shapes = await explorer.getChildren(folder);
        expect(say(shapes[0])).toBe('RunButton | AutoShape, runs DoIt | shape-edit-link-macro-delete');
        const pair = shapes.find((s) => s.label === 'Pair')!;
        // A group runs no macro, and a shape in it is not deleted on its own.
        expect(say(pair)).toBe('Pair | group | shape-edit-delete');
        const members = await explorer.getChildren(pair);
        expect(members.map((m) => explorer.getTreeItem(m).contextValue)).toEqual(['shape-edit-link', 'shape-edit-link']);
        expect(explorer.getParent(members[0])).toBe(pair);
        expect(explorer.getParent(folder)).toBe(sheet1);
        expect(explorer.getParent(sheet1)).toBe(sheets);
        // A click opens the editor.
        expect(explorer.getTreeItem(shapes[0]).command).toMatchObject({ command: 'xlide.editShape', arguments: [shapes[0]] });
        // Sheet2 has a shape and no module: its own row, with the same folder.
        const sheet2 = rows.find((row) => row.surface === 'Sheet2')!;
        expect(say(sheet2)).toBe('Sheet2 | shapeSurface-add');
        expect((await explorer.getChildren(sheet2)).map(say)).toEqual(['Shapes | shapeFolder-add']);
    });

    it('draw no Shapes folder for the workbook\'s own module, which has no sheet', async () => {
        const project = await projectOf('ShapesFixture.xlsm');
        const workbook = (await explorer.getChildren(project)).find((m) => m.moduleName === 'ThisWorkbook')!;
        expect((await explorer.getChildren(workbook)).some((row) => row.kind === 'shapes')).toBe(false);
    });

    it('list an ActiveX control without offering to edit it', async () => {
        // No ActiveX control can be inserted on the machine the fixtures were
        // made on, so the listing is stood in for.
        const bridge = {
            call: vi.fn(async (method: string) => {
                if (method === 'listModules') { return [{ name: 'Sheet1', type: 'document', documentType: 'worksheet' }]; }
                if (method === 'listSubs') { return []; }
                if (method === 'listWorkbookSheets') { return { sheets: [{ name: 'Data', codeName: 'Sheet1', kind: 'worksheet' }] }; }
                if (method === 'listShapes') {
                    return { surfaces: [{ surface: 'Data', codeName: 'Sheet1', shapes: [{ name: 'CommandButton1', kind: 'activeX', range: 'B8:C9' }] }] };
                }
                return { isPasswordProtected: false, isSigned: false };
            }),
        } as unknown as ConstructorParameters<typeof ProjectExplorer>[0];
        const standIn = new ProjectExplorer(bridge);
        vscodeMock.findFiles.mockResolvedValue([{ scheme: 'file', fsPath: 'C:\\work\\Book.xlsm' }]);
        const [project] = await standIn.getChildren();
        const [sheets] = await standIn.getChildren(project);
        const [sheet1] = await standIn.getChildren(sheets);
        expect(sheet1.label).toBe('Sheet1 (Data)');
        const [folder] = await standIn.getChildren(sheet1);
        const [control] = await standIn.getChildren(folder);
        const item = standIn.getTreeItem(control);
        expect(item).toMatchObject({ description: 'ActiveX control', contextValue: 'shape', command: undefined });
        expect(String(item.tooltip)).toMatch(/runs event procedures in its sheet's module/);
    });

    it('draw the Sheets folder again after XLIDE\'s own shape edit, which can move a sheet between its rows', async () => {
        const file = path.join(dir, 'SheetsFixture.xlsm');
        const project = await projectOf('SheetsFixture.xlsm');
        const { sheets, rows } = await sheetsOf(project);
        const bare = rows[2];
        expect((await explorer.getChildren(bare)).map((row) => row.label)).toEqual(['Trend', 'Later', 'Hidden']);

        editShape(file, 'Later', { action: 'add', type: 'rectangle', name: 'LaterBox', range: 'B2:C3' });
        vscodeMock.fired = [];
        explorer.refreshShapes(file, { shapesChanged: true });

        expect(vscodeMock.fired).toContain(sheets);
        const again = await explorer.getChildren(sheets);
        expect(again.map((row) => row.label)).toEqual(['Sheet1 (Budget)', 'Drawn', 'Later', 'Sheets With No Modules or Shapes']);
        expect((await explorer.getChildren(again[3])).map((row) => row.label)).toEqual(['Trend', 'Hidden']);
    });

    it('read only the sheet list on a change, and draw the Sheets folder again only when the sheets changed', async () => {
        const file = path.join(dir, 'SheetsFixture.xlsm');
        const project = await projectOf('SheetsFixture.xlsm');
        const { sheets } = await sheetsOf(project);
        const call = vi.spyOn(engine, 'call');
        vscodeMock.fired = [];

        // An auto-save on every keystroke: the sheet list is read again,
        // which is cheap; the shapes are not, and nothing is redrawn.
        explorer.refreshShapes(file);
        explorer.refreshShapes(file);
        await settle();
        expect(call.mock.calls.filter(([method]) => method === 'listShapes')).toHaveLength(0);
        expect(call.mock.calls.filter(([method]) => method === 'listWorkbookSheets').length).toBeGreaterThan(0);
        expect(vscodeMock.fired).toEqual([]);

        // Another workbook's bytes under the same name: other sheets, so the
        // Sheets folder is drawn again.
        fs.copyFileSync(path.join(FIXTURES, 'ShapesFixture.xlsm'), file);
        explorer.refreshShapes(file);
        await settle();
        expect(vscodeMock.fired).toContain(sheets);
        const rows = await explorer.getChildren(sheets);
        expect(rows.map((row) => row.label)).toEqual(['Sheet1 (Sheet1)', 'Sheet2']);
    });
});

describe('shapes in the explorer', () => {
    it('hangs a document\'s shapes under ThisDocument, by story', async () => {
        const project = await projectOf('ShapesArrangedFixture.docm');
        const rows = await explorer.getChildren(project);
        expect(rows.some((row) => row.kind === 'shapes')).toBe(false);
        const document = rows.find((m) => m.moduleName === 'ThisDocument')!;
        const [folder] = await explorer.getChildren(document);
        expect(say(folder)).toBe('Shapes | shapeFolder-add');
        const stories = await explorer.getChildren(folder);
        expect(stories[0]).toMatchObject({ kind: 'surface', surface: 'Document' });
        expect(explorer.getTreeItem(stories[0]).contextValue).toBe('shapeSurface-add');
        const shapes = await explorer.getChildren(stories[0]);
        expect(shapes.length).toBeGreaterThan(0);
        // Word cannot run a macro from a shape: no link action anywhere.
        expect(shapes.every((s) => !explorer.getTreeItem(s).contextValue?.includes('link'))).toBe(true);
    });

    it('lists a presentation\'s slides in a folder of their own, first under the project', async () => {
        const project = await projectOf('PowerPointShapesFixture.pptm');
        const [slides] = await explorer.getChildren(project);
        expect(say(slides)).toBe('Slides | shapeFolder');
        const [slide1, slide2] = await explorer.getChildren(slides);
        expect(say(slide1)).toBe('Slide 1 | 3 shapes | shapeSurface-add');
        expect(say(slide2)).toBe('Slide 2 | 2 shapes | shapeSurface-add');
        expect((await explorer.getChildren(slide1)).map((s) => s.label)).toEqual(['ClickMe', 'Caption', 'Badge']);
    });

    it('reads the file again when it changes, and says so on an empty sheet', async () => {
        const project = await projectOf('PowerPointShapesFixture.pptm', 'Deck.pptm');
        const [slides] = await explorer.getChildren(project);
        const [, slide2] = await explorer.getChildren(slides);
        expect((await explorer.getChildren(slide2)).map((s) => s.label)).toEqual(['SecondSlideShape', 'Pair']);

        editShape(path.join(dir, 'Deck.pptm'), 'Slide 2', { action: 'delete', name: 'Pair' });
        editShape(path.join(dir, 'Deck.pptm'), 'Slide 2', { action: 'delete', name: 'SecondSlideShape' });
        vscodeMock.fired = [];
        explorer.refreshShapes(path.join(dir, 'Deck.pptm'));
        expect(vscodeMock.fired).toContain(slides);
        const [, again] = await explorer.getChildren(slides);
        const rows = await explorer.getChildren(again);
        expect(rows.map(say)).toEqual(['No shapes | noShapes']);
    });

    it('lists a worksheet\'s shapes as the engine does', async () => {
        const project = await projectOf('ShapesFixture.xlsm');
        const { rows } = await sheetsOf(project);
        const sheet2 = rows.find((row) => row.surface === 'Sheet2')!;
        const [folder] = await explorer.getChildren(sheet2);
        const listed = listShapes(path.join(dir, 'ShapesFixture.xlsm'), 'Sheet2').surfaces[0].shapes.map((s) => s.name);
        expect((await explorer.getChildren(folder)).map((s) => s.label)).toEqual(listed);
    });
});
