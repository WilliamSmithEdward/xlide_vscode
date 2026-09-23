import * as vscode from 'vscode';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
    window: { setStatusBarMessage: vi.fn() },
}));
// The coordination with a running Office application has tests of its own;
// here the write simply runs.
vi.mock('../src/officeWriteCoordinator', () => ({
    runWriteWithHostCoordination: vi.fn((_filePath: string, write: () => Promise<unknown>) => write()),
}));

import { ProjectEngine } from '../src/projectEngine';
import { openShapeEditor, type ShapeEditorDeps } from '../src/shapeEditor';
import { editShape, listShapes } from '../src/vba/projectService';
import type { ShapeEditorModel, ShapeFormValues } from '../src/shapeEditorModel';
import type { ShapeInfo } from '../src/vba/shapes';

// The editor tab end to end: the page it opens on, and what each button does
// to a real file, with the webview itself stood in for.
const FIXTURES = path.join(__dirname, 'fixtures', 'binaries');

interface FakePanel {
    webview: { html: string; postMessage: ReturnType<typeof vi.fn> };
    reveal: ReturnType<typeof vi.fn>;
    dispose: ReturnType<typeof vi.fn>;
    send(message: unknown): Promise<void>;
}

function fakePanel(): FakePanel {
    let receive: ((message: unknown) => Promise<void>) | undefined;
    let disposed: (() => void) | undefined;
    const panel = {
        webview: {
            html: '',
            postMessage: vi.fn(async () => true),
            onDidReceiveMessage: vi.fn((handler: (message: unknown) => Promise<void>) => {
                receive = handler;
                return { dispose: vi.fn() };
            }),
        },
        reveal: vi.fn(),
        onDidDispose: vi.fn((handler: () => void) => {
            disposed = handler;
            return { dispose: vi.fn() };
        }),
        dispose: vi.fn(() => disposed?.()),
        send: (message: unknown) => receive!(message),
    };
    return panel as unknown as FakePanel;
}

let dir: string;
let deck: string;
let panel: FakePanel;
let deps: ShapeEditorDeps & { explorer: { refreshShapes: ReturnType<typeof vi.fn> } };
const context = { extensionUri: { fsPath: path.join(__dirname, '..') } } as unknown as vscode.ExtensionContext;

beforeEach(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-shape-editor-'));
    deck = path.join(dir, 'Deck.pptm');
    fs.copyFileSync(path.join(FIXTURES, 'PowerPointShapesFixture.pptm'), deck);
    panel = fakePanel();
    vi.mocked(vscode.window.createWebviewPanel).mockReset().mockReturnValue(panel as unknown as vscode.WebviewPanel);
    vi.mocked(vscode.window.showWarningMessage).mockReset();
    deps = {
        bridge: new ProjectEngine({} as never),
        explorer: { refreshShapes: vi.fn() } as never,
        out: { appendLine: vi.fn() } as never,
        goToMacro: vi.fn(async () => undefined),
    };
});

afterEach(() => {
    panel.dispose();
    fs.rmSync(dir, { recursive: true, force: true });
});

const shapesOn = (surface: string): ShapeInfo[] => listShapes(deck, surface).surfaces[0].shapes;
const shape = (surface: string, name: string) => shapesOn(surface).find((s) => s.name === name);

/** The model the page was drawn from, as its script holds it. */
function pageModel(): ShapeEditorModel {
    const json = /const model = (\{.*\});/.exec(panel.webview.html)?.[1];
    expect(json, 'the page carries its model').toBeDefined();
    return JSON.parse(json!) as ShapeEditorModel;
}

async function openOn(name: string): Promise<ShapeFormValues> {
    await openShapeEditor(deps, context, deck, { host: 'powerpoint', surface: 'Slide 1', shape: shape('Slide 1', name)! });
    return pageModel().values;
}

describe('the shape editor tab', () => {
    it('opens on what the file holds now, not on the tree\'s older listing', async () => {
        const listedEarlier = shape('Slide 1', 'ClickMe')!;
        editShape(deck, 'Slide 1', { action: 'update', name: 'ClickMe', text: 'Changed outside' });
        await openShapeEditor(deps, context, deck, { host: 'powerpoint', surface: 'Slide 1', shape: listedEarlier });

        const model = pageModel();
        expect(model.values.text).toBe('Changed outside');
        expect(model.macros.map((m) => m.macro)).toEqual(['SayHello', 'Unlinked']);
        expect(model.stackCount).toBe(3);
        expect(panel.webview.html).toContain('<title>ClickMe - Shape</title>');
    });

    it('saves what changed, refreshes the tree and closes', async () => {
        const values = await openOn('ClickMe');
        await panel.send({ type: 'save', values: { ...values, name: 'Go', fillType: 'solid', fillColor: '#00ff00', macro: 'Unlinked' } });

        expect(shape('Slide 1', 'Go')).toMatchObject({ macro: 'Unlinked', fill: { type: 'solid', color: '#00FF00' } });
        expect(deps.explorer.refreshShapes).toHaveBeenCalledWith(deck);
        expect(panel.dispose).toHaveBeenCalled();
    });

    it('closes without writing when nothing changed', async () => {
        const values = await openOn('Badge');
        const before = fs.readFileSync(deck);
        await panel.send({ type: 'save', values });

        expect(fs.readFileSync(deck).equals(before)).toBe(true);
        expect(panel.dispose).toHaveBeenCalled();
    });

    it('says what is wrong with the form, and leaves the file and the tab alone', async () => {
        const values = await openOn('Badge');
        const before = fs.readFileSync(deck);
        await panel.send({ type: 'save', values: { ...values, fillType: 'solid', fillColor: 'green', width: '-5' } });

        expect(panel.webview.postMessage).toHaveBeenCalledWith({
            type: 'invalid',
            errors: { fillColor: 'The fill color is a color as #RRGGBB, such as #FF0000.', width: 'The width is above zero.' },
        });
        expect(fs.readFileSync(deck).equals(before)).toBe(true);
        expect(panel.dispose).not.toHaveBeenCalled();
    });

    it('keeps the tab open with the reason when the application would refuse the change', async () => {
        const values = await openOn('Caption');
        await panel.send({ type: 'save', values: { ...values, name: 'Badge' } });

        expect(panel.webview.postMessage).toHaveBeenCalledWith({
            type: 'error',
            error: expect.stringMatching(/Slide 1 already has a shape named 'Badge'/),
        });
        expect(panel.dispose).not.toHaveBeenCalled();
    });

    it('deletes the shape once asked, and not before', async () => {
        await openOn('Caption');
        vi.mocked(vscode.window.showWarningMessage).mockResolvedValueOnce(undefined);
        await panel.send({ type: 'delete' });
        expect(shape('Slide 1', 'Caption')).toBeDefined();
        expect(panel.webview.postMessage).toHaveBeenCalledWith({ type: 'idle' });

        vi.mocked(vscode.window.showWarningMessage).mockResolvedValueOnce('Delete' as never);
        await panel.send({ type: 'delete' });
        expect(shape('Slide 1', 'Caption')).toBeUndefined();
        expect(panel.dispose).toHaveBeenCalled();
    });

    it('adds a shape where it was asked, with what the form gave', async () => {
        await openShapeEditor(deps, context, deck, { host: 'powerpoint', surface: 'Slide 2' });
        const values = pageModel().values;
        await panel.send({ type: 'save', values: { ...values, type: 'oval', name: 'Dot', left: '10', top: '20', macro: 'SayHello' } });

        expect(shape('Slide 2', 'Dot')).toMatchObject({ kind: 'shape', geometry: 'ellipse', left: 10, top: 20, macro: 'SayHello' });
    });

    it('brings forward the tab already open on a shape, rather than a second one', async () => {
        await openOn('Badge');
        await openOn('Badge');
        expect(vscode.window.createWebviewPanel).toHaveBeenCalledTimes(1);
        expect(panel.reveal).toHaveBeenCalled();
    });

    it('opens the Sub the form names', async () => {
        await openOn('ClickMe');
        await panel.send({ type: 'goToMacro', macro: 'Macros.SayHello' });
        expect(deps.goToMacro).toHaveBeenCalledWith(deck, 'Macros.SayHello');
    });
});
