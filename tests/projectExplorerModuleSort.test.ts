import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

const host = vi.hoisted(() => ({ compare: vi.fn(), findFiles: vi.fn() }));
vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
    workspace: { findFiles: host.findFiles, workspaceFolders: [{ uri: { fsPath: 'C:/work' } }] },
}));
vi.mock('../src/moduleDisplay', async (importOriginal) => {
    const real = await importOriginal<typeof import('../src/moduleDisplay')>();
    return { ...real, compareVbaModulesForTreeOrder: host.compare.mockImplementation(real.compareVbaModulesForTreeOrder) };
});
import { ProjectExplorer } from '../src/projectExplorer';
import { compareVbaModulesForTreeOrder } from '../src/moduleDisplay';

const BOOK = 'C:/work/App.vbp';
const modules = Object.freeze(Array.from({ length: 1000 }, (_, i) => Object.freeze({
    name: 'M' + String((i * 317) % 1000).padStart(4, '0'),
    type: ['standard', 'class', 'userform'][i % 3],
})));
const inputOrder = modules.map(module => module.name);
let explorers: ProjectExplorer[] = [];
beforeEach(() => {
    explorers = [];
    host.findFiles.mockResolvedValue([{ scheme: 'file', fsPath: BOOK }]);
});
afterEach(() => { for (const explorer of explorers) explorer.dispose(); });

async function measure(callers: number) {
    host.compare.mockClear();
    let release!: (value: typeof modules) => void;
    const pending = new Promise<typeof modules>(resolve => { release = resolve; });
    const call = vi.fn(() => pending);
    const explorer = new ProjectExplorer({ call } as unknown as ConstructorParameters<typeof ProjectExplorer>[0]);
    explorers.push(explorer);
    const [project] = await explorer.getChildren();
    const expansion = explorer.getChildren(project);
    const follows = Array.from({ length: callers - 1 }, () => explorer.resolveModuleNode(BOOK, 'M0317'));
    // Follow traverses the cached project root while expansion awaits listModules.
    await Promise.resolve();
    await Promise.resolve();
    release(modules);
    const rows = await expansion;
    for (const node of await Promise.all(follows)) {
        expect(node).toBe(rows.find(row => row.moduleName === 'M0317'));
    }
    expect(call).toHaveBeenCalledTimes(1);
    const comparisons = host.compare.mock.calls.length;
    expect(comparisons).toBeGreaterThan(0);
    expect(modules.map(module => module.name)).toEqual(inputOrder);
    const expected = [...modules].sort(compareVbaModulesForTreeOrder).map(module => module.name);
    expect(rows.map(row => row.moduleName)).toEqual(expected);
    host.compare.mockClear();
    const cached = await explorer.getChildren(project);
    expect(host.compare).not.toHaveBeenCalled();
    expect(call).toHaveBeenCalledTimes(1);
    for (let index = 0; index < rows.length; index++) expect(cached[index]).toBe(rows[index]);
    return comparisons;
}

describe('coalesced module sorting', () => {
    it.each([2, 8])('shares sorting across %i overlapping expansion/follow callers', async callers => {
        const single = await measure(1);
        const overlap = await measure(callers);
        expect(overlap).toBe(single);
    });
});
