import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
const host = vi.hoisted(() => ({ findFiles: vi.fn(), showErrorMessage: vi.fn() }));
vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
    workspace: { findFiles: host.findFiles, workspaceFolders: [{ uri: { fsPath: 'C:/work' } }] },
    window: { showErrorMessage: host.showErrorMessage },
}));
import { ProjectExplorer } from '../src/projectExplorer';
const BOOK = 'C:/work/Book.xlsm';
let explorers: ProjectExplorer[] = [];
function create(call: ReturnType<typeof vi.fn>) {
    const explorer = new ProjectExplorer({ call } as unknown as ConstructorParameters<typeof ProjectExplorer>[0]);
    explorers.push(explorer); return explorer;
}
beforeEach(() => { explorers = []; host.findFiles.mockResolvedValue([{ scheme: 'file', fsPath: BOOK }]); });
afterEach(() => explorers.forEach(explorer => explorer.dispose()));
const flush = async () => { for (let i = 0; i < 15; i++) { await Promise.resolve(); } };
describe('resolved module readiness', () => {
    it('shares sheet layout work across eight overlapping expansions', async () => {
        async function measure(count: number) {
            let reads = 0;
            const sheets = Array.from({ length: 1000 }, (_, i) => ({
                name: `Data${i}`, kind: 'worksheet', get codeName() { reads++; return `Sheet${i}`; },
            }));
            const call = vi.fn(async (method: string) => method === 'listModules'
                ? [{ name: 'Sheet1', type: 'document' }] : { sheets });
            const explorer = create(call), [project] = await explorer.getChildren();
            const rows = await Promise.all(Array.from({ length: count }, () => explorer.getChildren(project)));
            expect(rows.every(row => row[0] === rows[0][0])).toBe(true);
            expect(call.mock.calls.filter(([method]) => method === 'listWorkbookSheets')).toHaveLength(1);
            return reads;
        }
        expect(await measure(8)).toBe(await measure(1));
    });

    it('waits for a pending sheet layout before returning a cached module', async () => {
        let resolve!: (value: unknown) => void;
        const catalog = new Promise(yes => { resolve = yes; });
        const call = vi.fn((method: string) => method === 'listModules'
            ? Promise.resolve([{ name: 'Sheet1', type: 'document' }]) : catalog);
        const explorer = create(call), [project] = await explorer.getChildren();
        const drawing = explorer.getChildren(project);
        await flush();
        expect(explorer.getModuleNode(BOOK, 'Sheet1')).toBeDefined();
        let settled = false;
        const follow = explorer.resolveModuleNode(BOOK, 'Sheet1').then(node => { settled = true; return node; });
        await flush();
        expect(settled).toBe(false);
        resolve({ sheets: [{ name: 'Data', codeName: 'Sheet1', kind: 'worksheet' }] });
        await drawing;
        expect(explorer.getParent((await follow)!)).toMatchObject({ shapeFolder: 'sheets' });
    });

    it('starts a current render after sheet refresh while an old catalog remains pending', async () => {
        let resolve!: (value: unknown) => void, catalogCalls = 0;
        const old = new Promise(yes => { resolve = yes; });
        const call = vi.fn((method: string) => method === 'listModules'
            ? Promise.resolve([{ name: 'Sheet1', type: 'document' }])
            : ++catalogCalls === 1 ? old : Promise.resolve({ sheets: [{ name: 'New', codeName: 'Sheet1', kind: 'worksheet' }] }));
        const explorer = create(call), [project] = await explorer.getChildren();
        const drawing = explorer.getChildren(project);
        await flush();
        explorer.refreshShapes(BOOK);
        let settled = false;
        const fresh = explorer.getChildren(project).then(rows => { settled = true; return rows; });
        await flush();
        const completedBeforeOldRead = settled;
        resolve({ sheets: [{ name: 'Old', codeName: 'Sheet1', kind: 'worksheet' }] });
        await Promise.all([drawing, fresh]);
        expect(completedBeforeOldRead).toBe(true);
        expect(explorer.getModuleNode(BOOK, 'Sheet1')?.label).toBe('Sheet1 (New)');
        expect(catalogCalls).toBe(2);
    });

    it.each(['removed', 'failed'] as const)('does not resolve retained module objects when the new listing is %s', async outcome => {
        const call = vi.fn(async (method: string) => method === 'listModules'
            ? [{ name: 'M', type: 'standard' }] : { sheets: [] });
        const explorer = create(call), [project] = await explorer.getChildren();
        await explorer.getChildren(project);
        call.mockImplementation(async method => {
            if (method === 'listModules' && outcome === 'failed') { throw new Error('Workbook busy'); }
            return method === 'listModules' ? [] : { sheets: [] };
        });
        explorer.retryLoad({ kind: 'loadError', label: 'Retry', filePath: BOOK });
        expect(await explorer.resolveModuleNode(BOOK, 'M')).toBeUndefined();
        expect(call.mock.calls.filter(([method]) => method === 'listModules')).toHaveLength(2);
    });

    it('rebuilds the reveal path after an editor folder override is forgotten', async () => {
        const call = vi.fn(async (method: string) => method === 'listModules'
            ? [{ name: 'M', type: 'standard', folder: 'Saved' }] : { sheets: [] });
        const explorer = create(call); explorer.setView('folders');
        const module = await explorer.resolveModuleNode(BOOK, 'M');
        expect(module).toBeDefined();
        explorer.setModuleFolder(BOOK, 'M', 'Edited');
        explorer.forgetModuleFolder(BOOK, 'M');
        expect(await explorer.resolveModuleNode(BOOK, 'M')).toBe(module);
        expect(explorer.getParent(module!)?.folder).toBe('Saved');
        expect(call.mock.calls.filter(([method]) => method === 'listModules')).toHaveLength(2);
    });

    it('rebuilds an edited folder path without rereading cached modules', async () => {
        const call = vi.fn(async (method: string) => method === 'listModules'
            ? [{ name: 'M', type: 'standard', folder: 'Saved' }] : { sheets: [] });
        const explorer = create(call); explorer.setView('folders');
        const module = await explorer.resolveModuleNode(BOOK, 'M');
        explorer.setModuleFolder(BOOK, 'M', 'Edited');
        expect(await explorer.resolveModuleNode(BOOK, 'M')).toBe(module);
        expect(explorer.getParent(module!)?.folder).toBe('Edited');
        expect(call.mock.calls.filter(([method]) => method === 'listModules')).toHaveLength(1);
    });

    it('does not resolve cached modules or procedures after disposal', async () => {
        const call = vi.fn(async (method: string) => method === 'listModules'
            ? [{ name: 'M', type: 'standard' }] : method === 'listSubs'
                ? [{ name: 'Run', kind: 'Sub', line: 1 }] : { sheets: [] });
        const explorer = create(call);
        expect(await explorer.resolveProcedureNode(BOOK, 'M', 'Sub Run')).toBeDefined();
        const before = call.mock.calls.length;
        explorer.dispose();
        expect(await explorer.resolveModuleNode(BOOK, 'M')).toBeUndefined();
        expect(await explorer.resolveProcedureNode(BOOK, 'M', 'Sub Run')).toBeUndefined();
        expect(call).toHaveBeenCalledTimes(before);
    });
});
