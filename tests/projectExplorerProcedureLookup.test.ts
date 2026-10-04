import { afterEach, beforeEach, expect, it, vi } from 'vitest';
const host = vi.hoisted(() => ({ findFiles: vi.fn() }));
vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
    workspace: { findFiles: host.findFiles, workspaceFolders: [{ uri: { fsPath: 'C:/work' } }] },
}));
import { ProjectExplorer } from '../src/projectExplorer';
const BOOK = 'C:/work/App.vbp';
let explorers: ProjectExplorer[] = [];
function create(subs: () => Promise<Array<{name: string; kind: string; line: number}>>) {
    const call = vi.fn((method: string) => method === 'listModules'
        ? Promise.resolve([{ name: 'M', type: 'userform' }]) : method === 'listSubs'
        ? subs() : Promise.resolve({ isPasswordProtected: false, isSigned: false }));
    const explorer = new ProjectExplorer({ call } as unknown as ConstructorParameters<typeof ProjectExplorer>[0]);
    explorers.push(explorer); return { explorer, call };
}
beforeEach(() => { explorers = []; host.findFiles.mockResolvedValue([{ scheme: 'file', fsPath: BOOK }]); });
afterEach(() => { for (const explorer of explorers) explorer.dispose(); });

it('does not scan cached procedure labels on repeated follow lookups', async () => {
    const { explorer, call } = create(async () => Array.from({ length: 1000 }, (_, i) => ({ name: 'P' + i, kind: 'Sub', line: i + 1 })));
    const module = (await explorer.resolveModuleNode(BOOK, 'M'))!;
    const rows = await explorer.getChildren(module);
    let reads = 0;
    for (const row of rows) {
        const label = row.label;
        Object.defineProperty(row, 'label', { configurable: true, get: () => { reads++; return label; } });
    }
    for (let i = 0; i < 200; i++) expect(await explorer.resolveProcedureNode(BOOK, 'M', 'sub p999')).toBe(rows[1000]);
    expect(reads).toBe(0);
    expect(call.mock.calls.filter(([method]) => method === 'listSubs')).toHaveLength(1);
});

it('preserves the first case-insensitive duplicate, property kinds and designer exclusion', async () => {
    const { explorer } = create(async () => [
        { name: 'Same', kind: 'Sub', line: 1 }, { name: 'SAME', kind: 'Sub', line: 5 },
        { name: 'Name', kind: 'Property Get', line: 9 }, { name: 'Name', kind: 'Property Let', line: 12 },
    ]);
    const module = (await explorer.resolveModuleNode(BOOK, 'M'))!, rows = await explorer.getChildren(module);
    expect(explorer.getProcedureNode(BOOK, 'M', 'SUB SAME')).toBe(rows[1]);
    expect(explorer.getProcedureNode(BOOK, 'M', 'property get name')).toBe(rows[3]);
    expect(explorer.getProcedureNode(BOOK, 'M', 'property let name')).toBe(rows[4]);
    expect(explorer.getProcedureNode(BOOK, 'M', 'Designer')).toBeUndefined();
    expect(explorer.getProcedureNode(BOOK, 'M', 'Sub Absent')).toBeUndefined();
});

it.each(['module', 'project'] as const)('discards the lookup with %s refresh and retains new row identity', async scope => {
    let name = 'Old';
    const { explorer } = create(async () => [{ name, kind: 'Sub', line: 1 }]);
    const old = await explorer.resolveProcedureNode(BOOK, 'M', 'Sub Old');
    name = 'New';
    if (scope === 'module') explorer.refreshModuleSubs(BOOK, 'M'); else explorer.refresh();
    expect(explorer.getProcedureNode(BOOK, 'M', 'Sub Old')).toBeUndefined();
    const fresh = await explorer.resolveProcedureNode(BOOK, 'M', 'Sub New');
    expect(fresh).toBeDefined(); expect(fresh).not.toBe(old);
    expect(await explorer.resolveProcedureNode(BOOK, 'M', 'SUB NEW')).toBe(fresh);
    expect(explorer.getProcedureNode(BOOK, 'M', 'Sub Old')).toBeUndefined();
});
