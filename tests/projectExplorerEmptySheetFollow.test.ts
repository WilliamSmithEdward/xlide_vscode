import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
const host = vi.hoisted(() => ({ findFiles: vi.fn() }));
vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
    workspace: { findFiles: host.findFiles, workspaceFolders: [{ uri: { fsPath: 'C:/work' } }] },
}));
import { ProjectExplorer } from '../src/projectExplorer';
const BOOK = 'C:/work/Book.xlsm';
let explorers: ProjectExplorer[] = [];
function create(call = vi.fn(async (method: string) => method === 'listModules'
    ? [{ name: 'Sheet1', type: 'document', hasCode: false }]
    : method === 'listWorkbookSheets' ? { sheets: [{ name: 'Data', codeName: 'Sheet1', kind: 'worksheet' }] }
        : { surfaces: [] })) {
    const explorer = new ProjectExplorer({ call } as unknown as ConstructorParameters<typeof ProjectExplorer>[0]);
    explorers.push(explorer); return { explorer, call };
}
beforeEach(() => { explorers = []; host.findFiles.mockResolvedValue([{ scheme: 'file', fsPath: BOOK }]); });
afterEach(() => explorers.forEach(explorer => explorer.dispose()));
describe('following empty sheet modules', () => {
    it('builds the bare-sheet reveal path before resolving a cold module', async () => {
        const { explorer } = create();
        const module = await explorer.resolveModuleNode(BOOK, 'Sheet1');
        expect(module).toBeDefined();
        const bare = explorer.getParent(module!)!;
        expect(bare.shapeFolder).toBe('bareSheets');
        expect(await explorer.getChildren(bare)).toEqual([module]);
        const sheets = explorer.getParent(bare)!;
        expect(sheets.shapeFolder).toBe('sheets');
        expect(explorer.getParent(sheets)).toBe((await explorer.getChildren())[0]);
    });

    it('keeps an empty module with shapes directly under Sheets', async () => {
        const { explorer, call } = create();
        const base = call.getMockImplementation()!;
        call.mockImplementation(async method => method === 'listShapes'
            ? { surfaces: [{ surface: 'Data', codeName: 'Sheet1', shapes: [{ name: 'Box', kind: 'shape' }] }] }
            : base(method));
        const module = await explorer.resolveModuleNode(BOOK, 'Sheet1');
        expect(explorer.getParent(module!)?.shapeFolder).toBe('sheets');
    });

    it.each([false, true])('keeps repeated warm follows cached without scanning sheet metadata (shapes: %s)', async hasShapes => {
        let names = 0;
        const call = vi.fn(async (method: string) => method === 'listModules'
            ? [{ name: 'Sheet1', type: 'document', hasCode: false }]
            : method === 'listWorkbookSheets' ? { sheets: [{ get name() { names++; return 'Data'; }, codeName: 'Sheet1', kind: 'worksheet' }] }
                : { surfaces: hasShapes ? [{ surface: 'Data', codeName: 'Sheet1', shapes: [{ name: 'Box', kind: 'shape' }] }] : [] });
        const { explorer } = create(call);
        const module = await explorer.resolveModuleNode(BOOK, 'Sheet1');
        names = 0;
        for (let i = 0; i < 200; i++) { expect(await explorer.resolveModuleNode(BOOK, 'Sheet1')).toBe(module); }
        expect(names).toBe(0);
        expect(call.mock.calls.filter(([method]) => method === 'listShapes')).toHaveLength(1);
    });

    it('preserves a confirmed bare parent through cached project redraws', async () => {
        const { explorer, call } = create();
        const module = await explorer.resolveModuleNode(BOOK, 'Sheet1');
        const parent = explorer.getParent(module!)!;
        const [project] = await explorer.getChildren();
        await explorer.getChildren(project);
        expect(explorer.getParent(module!)).toBe(parent);
        expect(await explorer.resolveModuleNode(BOOK, 'Sheet1')).toBe(module);
        expect(call.mock.calls.filter(([method]) => method === 'listShapes')).toHaveLength(1);
    });

    it('reclassifies the parent when a sheet gains or loses its last shape', async () => {
        const { explorer, call } = create();
        const base = call.getMockImplementation()!;
        const module = await explorer.resolveModuleNode(BOOK, 'Sheet1');
        expect(explorer.getParent(module!)?.shapeFolder).toBe('bareSheets');
        call.mockImplementation(async method => method === 'listShapes'
            ? { surfaces: [{ surface: 'Data', codeName: 'Sheet1', shapes: [{ name: 'Box', kind: 'shape' }] }] }
            : base(method));
        explorer.refreshShapes(BOOK, { shapesChanged: true });
        expect(await explorer.resolveModuleNode(BOOK, 'Sheet1')).toBe(module);
        expect(explorer.getParent(module!)?.shapeFolder).toBe('sheets');
        call.mockImplementation(base);
        explorer.refreshShapes(BOOK, { shapesChanged: true });
        expect(await explorer.resolveModuleNode(BOOK, 'Sheet1')).toBe(module);
        expect(explorer.getParent(module!)?.shapeFolder).toBe('bareSheets');
        expect(call.mock.calls.filter(([method]) => method === 'listShapes')).toHaveLength(3);
        expect(call.mock.calls.filter(([method]) => method === 'listModules')).toHaveLength(1);
    });

    it('does not read shapes when following an empty standard module', async () => {
        const call = vi.fn(async (method: string) => method === 'listModules'
            ? [{ name: 'M', type: 'standard', hasCode: false }] : { sheets: [] });
        const { explorer } = create(call);
        expect(await explorer.resolveModuleNode(BOOK, 'M')).toBeDefined();
        expect(call.mock.calls.some(([method]) => method === 'listShapes')).toBe(false);
    });
});
