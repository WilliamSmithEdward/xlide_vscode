import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
const host = vi.hoisted(() => ({ findFiles: vi.fn() }));
vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
    workspace: { findFiles: host.findFiles, workspaceFolders: [{ uri: { fsPath: 'C:/work' } }] },
}));
import { ProjectExplorer } from '../src/projectExplorer';
import { keepAgentChange, presentAgentModuleWrite } from '../src/xlideAgentDiff';
const BOOK = 'C:/work/App.vbp', OTHER = 'C:/work/Other.vbp';
let explorers: ProjectExplorer[] = [], reviews: Array<[string, string]> = [];
beforeEach(() => { explorers = []; reviews = []; });
afterEach(() => {
    for (const [path, name] of reviews) keepAgentChange(path, name);
    for (const explorer of explorers) explorer.dispose();
});
async function pending(name: string, path = BOOK) {
    reviews.push([path, name]);
    await presentAgentModuleWrite(path, name, { before: 'Sub Old()\nEnd Sub', beforeExisted: true, after: 'Sub New()\nEnd Sub' });
}
async function create(modules: Array<{ name: string; type: string; folder?: string }>, other = false) {
    host.findFiles.mockResolvedValue((other ? [BOOK, OTHER] : [BOOK]).map(fsPath => ({ scheme: 'file', fsPath })));
    const explorer = new ProjectExplorer({ call: vi.fn(async () => modules) } as unknown as ConstructorParameters<typeof ProjectExplorer>[0]);
    explorers.push(explorer); explorer.setView('folders');
    const projects = await explorer.getChildren();
    for (const project of projects) await explorer.getChildren(project);
    return explorer;
}
function countPaths(explorer: ProjectExplorer, names: string[], path = BOOK) {
    let reads = 0;
    for (const name of names) Object.defineProperty(explorer.getModuleNode(path, name)!, 'filePath', {
        configurable: true, get() { reads++; return path; },
    });
    return () => reads;
}
function badge(explorer: ProjectExplorer, folder: string, path = BOOK) {
    return explorer.getTreeItem(explorer.getFolderNode(path, folder)!);
}
const modules = Array.from({ length: 1000 }, (_, i) => ({
    name: 'M' + String(i).padStart(4, '0'), type: 'standard', folder: 'F' + String(i).padStart(4, '0'),
}));

describe('folder review lookups', () => {
    it('visits only the pending module when rendering 200 unrelated folders', async () => {
        const explorer = await create(modules);
        await pending('m0999');
        const reads = countPaths(explorer, modules.map(module => module.name));
        for (let i = 0; i < 200; i++) expect(badge(explorer, modules[i].folder).description).not.toContain('agent edit');
        expect(reads()).toBeLessThanOrEqual(200);
        expect(badge(explorer, 'F0999').description).toContain('agent edit');
    });

    it('does not scan loaded rows for an agent refresh of an unknown module', async () => {
        const explorer = await create(modules), reads = countPaths(explorer, modules.map(module => module.name));
        explorer.refreshAgentReviewMarks(BOOK, 'NotLoaded');
        expect(reads()).toBe(0);
    });

    it('keeps the no-review path free of module reads', async () => {
        const explorer = await create(modules), reads = countPaths(explorer, modules.map(module => module.name));
        for (let i = 0; i < 200; i++) expect(badge(explorer, modules[i].folder).description).not.toContain('agent edit');
        expect(reads()).toBe(0);
    });

    it('isolates projects with identical module names and skips unknown review identities', async () => {
        const explorer = await create([{ name: 'M', type: 'standard', folder: 'Shared' }], true);
        await pending('Missing'); await pending('m', OTHER);
        expect(badge(explorer, 'Shared').description).not.toContain('agent edit');
        expect(badge(explorer, 'Shared', OTHER).description).toContain('agent edit');
    });

    it('tracks nested folders, editor moves and review completion with normalized module names', async () => {
        const explorer = await create([{ name: 'Ledger', type: 'standard', folder: 'Accounts.Ledger' }]);
        const initial = badge(explorer, 'Accounts').id;
        await pending('LEDGER');
        expect(badge(explorer, 'Accounts').description).toContain('agent edit');
        expect(badge(explorer, 'Accounts.Ledger').description).toContain('agent edit');
        expect(badge(explorer, 'Accounts').id).toBe(initial);
        explorer.setModuleFolder(BOOK, 'ledger', 'Moved.Deep');
        const [project] = await explorer.getChildren(); await explorer.getChildren(project);
        expect(badge(explorer, 'Accounts').description).not.toContain('agent edit');
        expect(badge(explorer, 'Moved').description).toContain('agent edit');
        const moved = badge(explorer, 'Moved').id;
        keepAgentChange(BOOK, 'Ledger');
        expect(badge(explorer, 'Moved').description).not.toContain('agent edit');
        expect(badge(explorer, 'Moved').id).toBe(moved);
    });
});
