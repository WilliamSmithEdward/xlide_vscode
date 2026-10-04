import { beforeEach, describe, expect, it, vi } from 'vitest';
const handlers = vi.hoisted(() => new Map<string, (node: unknown) => Promise<void>>());
vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
    commands: { registerCommand: vi.fn((name: string, handler: (node: unknown) => Promise<void>) => {
        handlers.set(name, handler); return { dispose: vi.fn() };
    }) },
    window: { setStatusBarMessage: vi.fn(), showQuickPick: vi.fn() },
}));
vi.mock('../src/shapeEditor', () => ({ openShapeEditor: vi.fn(), writeShapeEdit: vi.fn(async () => 'Box') }));
import * as vscode from 'vscode';
import { ShapeRows } from '../src/shapeRows';
import { registerShapeCommands } from '../src/commands/shapeCommands';
import { writeShapeEdit } from '../src/shapeEditor';
import type { ProjectEngine } from '../src/projectEngine';
import type { XlideNode } from '../src/projectExplorer';

const BOOK = 'C:/work/Book.xlsm';
const folder: XlideNode = { kind: 'shapes', shapeFolder: 'surface', surface: 'Data', label: 'Shapes', filePath: BOOK };
function deferred<T>() {
    let resolve!: (value: T) => void;
    return { promise: new Promise<T>(yes => { resolve = yes; }), resolve: (value: T) => resolve(value) };
}
async function setup() {
    const bridge = { call: vi.fn(async (method: string) => method === 'shapeMacros'
        ? { macros: [{ macro: 'Helpers.Run', module: 'Helpers', proc: 'Run' }] }
        : { surfaces: [{ surface: 'Data', shapes: [{ name: 'Box', kind: 'shape', macro: 'Old' }] }] }) };
    const tree = new ShapeRows(bridge as unknown as ProjectEngine, () => {});
    const [node] = await tree.children(folder, async () => []);
    registerShapeCommands({ bridge, explorer: { shapeContextOf: (row: XlideNode) => tree.contextOf(row) },
        context: {}, out: { appendLine: vi.fn() } } as never);
    return { tree, node, bridge, run: (command: string) => handlers.get(command)!(node) };
}
beforeEach(() => {
    vi.clearAllMocks(); handlers.clear();
    vi.mocked(vscode.window.showQuickPick).mockReset();
    vi.mocked(vscode.window.showWarningMessage).mockReset();
});

describe('shape command lifetime', () => {
    it('does not reuse an earlier confirmation after the same row is redrawn', async () => {
        const { tree, node, run } = await setup();
        const original = tree.contextOf(node);
        const confirmation = deferred<never>();
        vi.mocked(vscode.window.showWarningMessage).mockReturnValueOnce(confirmation.promise);
        try {
            const pending = run('xlide.deleteShape');
            tree.refresh(BOOK);
            await tree.children(node, async () => []);
            expect(tree.contextOf(node)).toBeDefined();
            expect(tree.contextOf(node)).not.toBe(original);
            confirmation.resolve('Delete' as never);
            await pending;
            expect(writeShapeEdit).not.toHaveBeenCalled();
        } finally { tree.dispose(); }
    });

    it.each(['refresh', 'dispose'] as const)('does not delete after %s during confirmation', async change => {
        const { tree, run } = await setup();
        const confirmation = deferred<never>();
        vi.mocked(vscode.window.showWarningMessage).mockReturnValueOnce(confirmation.promise);
        try {
            const pending = run('xlide.deleteShape');
            if (change === 'refresh') { tree.refresh(BOOK); } else { tree.dispose(); }
            confirmation.resolve('Delete' as never);
            await pending;
            expect(writeShapeEdit).not.toHaveBeenCalled();
        } finally { tree.dispose(); }
    });

    it('does not offer macros if the row changes during their read', async () => {
        const { tree, bridge, run } = await setup();
        const macros = deferred<{ macros: Array<{ macro: string; module: string; proc: string }> }>();
        bridge.call.mockReturnValueOnce(macros.promise as never);
        try {
            const pending = run('xlide.linkShapeMacro');
            tree.refresh(BOOK);
            macros.resolve({ macros: [{ macro: 'Helpers.Run', module: 'Helpers', proc: 'Run' }] });
            await pending;
            expect(vscode.window.showQuickPick).not.toHaveBeenCalled();
            expect(writeShapeEdit).not.toHaveBeenCalled();
        } finally { tree.dispose(); }
    });

    it('does not link a macro after a refresh while the picker is open', async () => {
        const { tree, run } = await setup();
        const pick = deferred<never>();
        vi.mocked(vscode.window.showQuickPick).mockImplementationOnce(() => {
            tree.refresh(BOOK); return pick.promise;
        });
        try {
            const pending = run('xlide.linkShapeMacro');
            pick.resolve({ macro: 'Helpers.Run' } as never);
            await pending;
            expect(vscode.window.showQuickPick).toHaveBeenCalledTimes(1);
            expect(writeShapeEdit).not.toHaveBeenCalled();
        } finally { tree.dispose(); }
    });

    it.each(['xlide.deleteShape', 'xlide.linkShapeMacro', 'xlide.unlinkShapeMacro'])(
        'ignores a retained stale row when invoking %s', async command => {
            const { tree, run } = await setup();
            try {
                tree.refresh(BOOK);
                await run(command);
                expect(writeShapeEdit).not.toHaveBeenCalled();
                expect(vscode.window.showWarningMessage).not.toHaveBeenCalled();
                expect(vscode.window.showQuickPick).not.toHaveBeenCalled();
            } finally { tree.dispose(); }
        });

    it.each(['xlide.deleteShape', 'xlide.linkShapeMacro'])(
        'still writes a current row through %s after another project refreshes', async command => {
            const { tree, run } = await setup();
            vi.mocked(vscode.window.showWarningMessage).mockResolvedValueOnce('Delete' as never);
            vi.mocked(vscode.window.showQuickPick).mockResolvedValueOnce({ macro: 'Helpers.Run' } as never);
            try {
                tree.refresh('C:/work/Other.xlsm');
                await run(command);
                expect(writeShapeEdit).toHaveBeenCalledTimes(1);
                expect(writeShapeEdit).toHaveBeenCalledWith(expect.anything(), BOOK, 'Data',
                    expect.objectContaining({ name: 'Box' }), command);
            } finally { tree.dispose(); }
        });
});
