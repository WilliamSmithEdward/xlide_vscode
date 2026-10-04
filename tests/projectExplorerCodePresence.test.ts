import { afterEach, describe, expect, it, vi } from 'vitest';
const host = vi.hoisted(() => ({ findFiles: vi.fn() }));
vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
    workspace: { findFiles: host.findFiles, workspaceFolders: [{ uri: { fsPath: '/work' } }] },
}));
import { ProjectExplorer } from '../src/projectExplorer';
const BOOK = '/work/Book.xlsm';
const explorers: ProjectExplorer[] = [];
afterEach(() => { for (const explorer of explorers.splice(0)) { explorer.dispose(); } });
function fixture(savedCode = false, hasShapes = false, type = 'document') {
    host.findFiles.mockResolvedValue([{ scheme: 'file', fsPath: BOOK }]);
    const saved = { hasCode: savedCode };
    const call = vi.fn(async (method: string) => method === 'listModules'
        ? [{ name: 'Sheet1', type, hasCode: saved.hasCode }]
        : method === 'listWorkbookSheets' ? { sheets: type === 'document'
            ? [{ name: 'Data', codeName: 'Sheet1', kind: 'worksheet' }] : [] }
            : { surfaces: hasShapes ? [{ surface: 'Data', codeName: 'Sheet1', shapes: [{ name: 'Box', kind: 'shape' }] }] : [] });
    const explorer = new ProjectExplorer({ call } as unknown as ConstructorParameters<typeof ProjectExplorer>[0]);
    explorers.push(explorer);
    return { explorer, call, saved };
}

describe('sheet module code presence from the editor', () => {
    it('moves an edited empty sheet out of the bare folder without rereading shapes or modules', async () => {
        const { explorer, call } = fixture();
        const module = (await explorer.resolveModuleNode(BOOK, 'Sheet1'))!;
        const bare = explorer.getParent(module)!;
        expect(bare.shapeFolder).toBe('bareSheets');
        const before = call.mock.calls.length;
        explorer.setModuleCodePresence(BOOK, 'Sheet1', true);
        expect(await explorer.resolveModuleNode(BOOK, 'Sheet1')).toBe(module);
        expect(explorer.getParent(module)?.shapeFolder).toBe('sheets');
        expect(await explorer.getChildren(bare)).toEqual([]);
        expect(call.mock.calls.length).toBe(before);
    });

    it('moves a sheet back into the bare folder when the editor removes its code', async () => {
        const { explorer, call } = fixture(true);
        const module = (await explorer.resolveModuleNode(BOOK, 'Sheet1'))!;
        explorer.setModuleCodePresence(BOOK, 'Sheet1', false);
        expect(await explorer.resolveModuleNode(BOOK, 'Sheet1')).toBe(module);
        expect(explorer.getParent(module)?.shapeFolder).toBe('bareSheets');
        expect(call.mock.calls.filter(([method]) => method === 'listModules')).toHaveLength(1);
    });

    it('keeps a sheet with shapes under Sheets after its code is removed', async () => {
        const { explorer } = fixture(true, true);
        const module = (await explorer.resolveModuleNode(BOOK, 'Sheet1'))!;
        explorer.setModuleCodePresence(BOOK, 'Sheet1', false);
        expect(await explorer.resolveModuleNode(BOOK, 'Sheet1')).toBe(module);
        expect(explorer.getParent(module)?.shapeFolder).toBe('sheets');
    });

    it('keeps an editor override through a cold load and refresh, then drops it on close', async () => {
        const { explorer } = fixture();
        explorer.setModuleCodePresence(BOOK, 'Sheet1', true);
        expect((await explorer.resolveModuleNode(BOOK, 'Sheet1'))?.hasCode).toBe(true);
        explorer.refresh();
        const module = (await explorer.resolveModuleNode(BOOK, 'Sheet1'))!;
        expect(module.hasCode).toBe(true);
        explorer.forgetModuleFolder(BOOK, 'Sheet1');
        expect(await explorer.resolveModuleNode(BOOK, 'Sheet1')).toBe(module);
        expect(module.hasCode).toBe(false);
        expect(explorer.getParent(module)?.shapeFolder).toBe('bareSheets');
    });

    it('does not redraw sheets on edits that keep the same code presence', async () => {
        const { explorer, call } = fixture();
        await explorer.resolveModuleNode(BOOK, 'Sheet1');
        const changed = vi.fn(), replaced = vi.fn();
        explorer.onDidChangeTreeData(changed); explorer.onDidReplaceRows(replaced);
        for (let i = 0; i < 100; i++) { explorer.setModuleCodePresence(BOOK, 'Sheet1', true); }
        expect(changed).toHaveBeenCalledTimes(1);
        expect(replaced).toHaveBeenCalledTimes(1);
        expect(call.mock.calls.filter(([method]) => method === 'listShapes')).toHaveLength(1);
    });

    it('does not apply code-placement overrides to a standard module', async () => {
        const { explorer } = fixture(false, false, 'standard');
        const module = (await explorer.resolveModuleNode(BOOK, 'Sheet1'))!;
        const changed = vi.fn(); explorer.onDidChangeTreeData(changed);
        explorer.setModuleCodePresence(BOOK, 'Sheet1', true);
        expect(module.hasCode).toBe(false);
        expect(changed).not.toHaveBeenCalled();
    });

    it('updates saved metadata without keeping an editor override across refresh', async () => {
        const { explorer, saved } = fixture();
        const module = (await explorer.resolveModuleNode(BOOK, 'Sheet1'))!;
        saved.hasCode = true;
        explorer.setModuleCodePresence(BOOK, 'Sheet1', true, { fromEditor: false });
        expect(await explorer.resolveModuleNode(BOOK, 'Sheet1')).toBe(module);
        expect(explorer.getParent(module)?.shapeFolder).toBe('sheets');
        saved.hasCode = false;
        explorer.refresh();
        expect((await explorer.resolveModuleNode(BOOK, 'Sheet1'))?.hasCode).toBe(false);
    });

    it('keeps unsaved code presence when the saved index reports a different value', async () => {
        const { explorer } = fixture();
        const module = (await explorer.resolveModuleNode(BOOK, 'Sheet1'))!;
        explorer.setModuleCodePresence(BOOK, 'Sheet1', true);
        explorer.setModuleCodePresence(BOOK, 'Sheet1', false, { fromEditor: false });
        expect(module.hasCode).toBe(true);
        expect(await explorer.resolveModuleNode(BOOK, 'Sheet1')).toBe(module);
        expect(explorer.getParent(module)?.shapeFolder).toBe('sheets');
    });
});
