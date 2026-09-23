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

// The shape rows, drawn from files the applications saved and read by the
// real engine: where each host's shapes hang, what each row offers, and that
// a changed file shows through.
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

describe('shapes in the explorer', () => {
    it('hangs a worksheet\'s shapes under its module, above its procedures', async () => {
        const project = await projectOf('ShapesFixture.xlsm');
        const modules = await explorer.getChildren(project);
        const sheet1 = modules.find((m) => m.moduleName === 'Sheet1')!;
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
        // A click opens the editor.
        expect(explorer.getTreeItem(shapes[0]).command).toMatchObject({ command: 'xlide.editShape', arguments: [shapes[0]] });
    });

    it('draws no Shapes folder for the workbook\'s own module, which has no sheet', async () => {
        const project = await projectOf('ShapesFixture.xlsm');
        const workbook = (await explorer.getChildren(project)).find((m) => m.moduleName === 'ThisWorkbook')!;
        expect((await explorer.getChildren(workbook)).some((row) => row.kind === 'shapes')).toBe(false);
    });

    it('lists a worksheet with no module in a folder of its own, once its shapes are read', async () => {
        const project = await projectOf('ShapesFixture.xlsm');
        const first = await explorer.getChildren(project);
        expect(first.some((row) => row.kind === 'shapes')).toBe(false);
        await settle();
        // The read landed and asked for the project row to be drawn again.
        expect(vscodeMock.fired).toContain(project);
        const [sheets] = await explorer.getChildren(project);
        expect(say(sheets)).toBe('Sheets With No Module | 1 sheet | shapeFolder');
        const [sheet2] = await explorer.getChildren(sheets);
        expect(sheet2).toMatchObject({ kind: 'surface', surface: 'Sheet2' });
        const listed = listShapes(path.join(dir, 'ShapesFixture.xlsm'), 'Sheet2').surfaces[0].shapes.map((s) => s.name);
        const rows = (await explorer.getChildren(sheet2)).map((s) => s.label);
        expect(rows).toEqual(listed.length > 0 ? listed : ['No shapes']);
    });

    it('lists an ActiveX control without offering to edit it', async () => {
        // No ActiveX control can be inserted on the machine the fixtures were
        // made on, so the listing is stood in for.
        const bridge = {
            call: vi.fn(async (method: string) => {
                if (method === 'listModules') { return [{ name: 'Sheet1', type: 'document', documentType: 'worksheet' }]; }
                if (method === 'listSubs') { return []; }
                if (method === 'listShapes') {
                    return { surfaces: [{ surface: 'Data', codeName: 'Sheet1', shapes: [{ name: 'CommandButton1', kind: 'activeX', range: 'B8:C9' }] }] };
                }
                return { isPasswordProtected: false, isSigned: false };
            }),
        } as unknown as ConstructorParameters<typeof ProjectExplorer>[0];
        const standIn = new ProjectExplorer(bridge);
        vscodeMock.findFiles.mockResolvedValue([{ scheme: 'file', fsPath: 'C:\\work\\Book.xlsm' }]);
        const [project] = await standIn.getChildren();
        const [sheet1] = await standIn.getChildren(project);
        const [folder] = await standIn.getChildren(sheet1);
        const [control] = await standIn.getChildren(folder);
        const item = standIn.getTreeItem(control);
        expect(item).toMatchObject({ description: 'ActiveX control', contextValue: 'shape', command: undefined });
        expect(String(item.tooltip)).toMatch(/runs event procedures in its sheet's module/);
    });

    it('hangs a document\'s shapes under ThisDocument, by story', async () => {
        const project = await projectOf('WordShapesFixture.docm');
        const thisDocument = (await explorer.getChildren(project)).find((m) => m.moduleName === 'ThisDocument')!;
        const [folder] = await explorer.getChildren(thisDocument);
        expect(say(folder)).toBe('Shapes | shapeFolder-add');
        const stories = await explorer.getChildren(folder);
        // The body always, and any other story that holds a shape.
        expect(stories.map(say)).toEqual(['Document | 3 shapes | shapeSurface-add', 'Header | 1 shape | shapeSurface-add']);
        const body = await explorer.getChildren(stories[0]);
        // Word runs no macro from a shape, so none offers a link.
        expect(body.map(say)).toEqual([
            'InlineOval | AutoShape | shape-edit-delete',
            'AnchoredBox | AutoShape | shape-edit-delete',
            'Board | drawing canvas | shape-edit-delete',
        ]);
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

    it('costs nothing on a change while no shape row is open, and redraws the workbook only when its sheets change', async () => {
        const project = await projectOf('ShapesFixture.xlsm');
        await explorer.getChildren(project);
        await settle();
        await explorer.getChildren(project);
        const call = vi.spyOn(engine, 'call');
        vscodeMock.fired = [];

        // An auto-save on every keystroke: nothing read, nothing redrawn.
        explorer.refreshShapes(path.join(dir, 'ShapesFixture.xlsm'));
        explorer.refreshShapes(path.join(dir, 'ShapesFixture.xlsm'));
        expect(call).not.toHaveBeenCalled();
        expect(vscodeMock.fired).toEqual([]);

        // Drawn again, the row reads the file behind the scenes, and keeps
        // itself as it is when the sheets with no module are the same.
        await explorer.getChildren(project);
        await settle();
        expect(call.mock.calls.filter(([method]) => method === 'listShapes')).toHaveLength(1);
        expect(vscodeMock.fired).toEqual([]);
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
});
