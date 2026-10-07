import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
const host = vi.hoisted(() => ({ findFiles: vi.fn() }));
vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
    workspace: { findFiles: host.findFiles, workspaceFolders: [{ uri: { fsPath: 'C:/work' } }] },
}));
import * as vscode from 'vscode';
import { ProjectExplorer, type XlideNode } from '../src/projectExplorer';
const BOOK = 'C:/work/Book.xlsm';
const LOCKED = { isPasswordProtected: true, isSigned: true };
const OPEN = { isPasswordProtected: false, isSigned: false };
function deferred<T>() {
    let resolve!: (value: T) => void, reject!: (error: unknown) => void;
    const promise = new Promise<T>((yes, no) => { resolve = yes; reject = no; });
    return { promise, resolve, reject };
}
let explorers: ProjectExplorer[] = [];
function create(probe: () => Promise<typeof OPEN>) {
    const appendLine = vi.fn();
    const call = vi.fn((method: string) => {
        if (method === 'getProtectionInfo') return probe();
        if (method === 'listModules') return Promise.resolve([{ name: 'M', type: 'standard' }]);
        if (method === 'listWorkbookSheets') return Promise.resolve({ sheets: [] });
        if (method === 'listShapes') return Promise.resolve({ surfaces: [] });
        return Promise.resolve([]);
    });
    const explorer = new ProjectExplorer({ call } as unknown as ConstructorParameters<typeof ProjectExplorer>[0],
        { appendLine } as unknown as ConstructorParameters<typeof ProjectExplorer>[1], undefined, vscode.Uri.file('C:/extension'));
    explorers.push(explorer);
    const events: Array<XlideNode | undefined | null | void> = [];
    explorer.onDidChangeTreeData((node) => events.push(node));
    const render = async () => { const [project] = await explorer.getChildren(); await explorer.getChildren(project); return project; };
    return { explorer, appendLine, events, render };
}
const idle = () => vi.advanceTimersByTimeAsync(2000);
const flush = () => vi.advanceTimersByTimeAsync(0);
beforeEach(() => {
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout'] });
    host.findFiles.mockReset();
    host.findFiles.mockResolvedValue([{ scheme: 'file', fsPath: BOOK }]);
    explorers = [];
});
afterEach(() => { for (const explorer of explorers) explorer.dispose(); vi.useRealTimers(); });

describe('protection probe ownership', () => {
    it('ignores an old success arriving after refreshed badges are complete', async () => {
        const old = deferred<typeof OPEN>();
        const probe = vi.fn().mockReturnValueOnce(old.promise).mockResolvedValue(LOCKED);
        const { explorer, events, render } = create(probe);
        await render(); await idle();
        explorer.refresh(); const current = await render(); await idle();
        expect(current).toMatchObject(LOCKED);
        const badges = explorer.getTreeItem(current).description;
        expect(badges).toContain('[locked, signed]');
        events.length = 0;
        old.resolve(OPEN); await flush();
        expect(events).toEqual([]);
        expect(current).toMatchObject(LOCKED);
        expect(explorer.getTreeItem(current).description).toBe(badges);
        await explorer.getChildren(current); await idle();
        expect(probe).toHaveBeenCalledTimes(2);
    });

    it('keeps badges unknown until the current pending probe completes', async () => {
        const old = deferred<typeof OPEN>(), fresh = deferred<typeof OPEN>();
        const probe = vi.fn().mockReturnValueOnce(old.promise).mockReturnValueOnce(fresh.promise);
        const { explorer, events, render } = create(probe);
        await render(); await idle(); explorer.refresh(); const current = await render(); await idle();
        events.length = 0;
        try {
            old.resolve(OPEN); await flush();
            expect(current.isPasswordProtected).toBeUndefined();
            expect(current.isSigned).toBeUndefined();
            expect(events).toEqual([]);
        } finally { fresh.resolve(LOCKED); await flush(); }
        expect(current).toMatchObject(LOCKED);
    });

    it('does not let an old failure delete the newer pending load or start a duplicate', async () => {
        const old = deferred<typeof OPEN>(), fresh = deferred<typeof OPEN>();
        const probe = vi.fn().mockReturnValueOnce(old.promise).mockReturnValueOnce(fresh.promise).mockResolvedValue(OPEN);
        const { explorer, appendLine, render } = create(probe);
        await render(); await idle(); explorer.refresh(); const current = await render(); await idle();
        try {
            old.reject(new Error('obsolete probe')); await flush();
            await explorer.getChildren(current); await idle();
            expect(probe).toHaveBeenCalledTimes(2);
            expect(appendLine).not.toHaveBeenCalled();
        } finally { fresh.resolve(LOCKED); await flush(); }
        expect(current).toMatchObject(LOCKED);
    });

    it('logs current failures and permits another idle probe', async () => {
        const probe = vi.fn().mockRejectedValueOnce(new Error('current probe')).mockResolvedValue(LOCKED);
        const { explorer, appendLine, render } = create(probe);
        const current = await render(); await idle();
        expect(appendLine).toHaveBeenCalledTimes(1);
        expect(appendLine).toHaveBeenCalledWith(expect.stringContaining('current probe'));
        await explorer.getChildren(current); await idle();
        expect(probe).toHaveBeenCalledTimes(2);
        expect(current).toMatchObject(LOCKED);
    });

    it('permits retry after a synchronous bridge failure without retaining a settled load', async () => {
        const probe = vi.fn().mockImplementationOnce(() => { throw new Error('synchronous probe'); }).mockResolvedValue(LOCKED);
        const { explorer, appendLine, render } = create(probe);
        const current = await render(); await idle();
        expect(appendLine).toHaveBeenCalledTimes(1);
        await explorer.getChildren(current); await idle();
        expect(probe).toHaveBeenCalledTimes(2);
        expect(current).toMatchObject(LOCKED);
    });

    it('coalesces repeated rendering while pending and caches a negative result', async () => {
        const pending = deferred<typeof OPEN>(), probe = vi.fn().mockReturnValue(pending.promise);
        const { explorer, render } = create(probe);
        const current = await render(); await idle();
        for (let i = 0; i < 100; i++) await explorer.getChildren(current);
        await idle(); expect(probe).toHaveBeenCalledTimes(1);
        pending.resolve(OPEN); await flush();
        for (let i = 0; i < 100; i++) await explorer.getChildren(current);
        await idle(); expect(probe).toHaveBeenCalledTimes(1);
        expect(current).toMatchObject(OPEN);
    });

    it('cancels an idle timer on refresh before it starts backend work', async () => {
        const probe = vi.fn().mockResolvedValue(OPEN);
        const { explorer, render } = create(probe);
        await render(); explorer.refresh(); const current = await render();
        await idle(); expect(probe).toHaveBeenCalledTimes(1); expect(current).toMatchObject(OPEN);
    });

    it('keeps probe ownership when only module folder data is invalidated', async () => {
        const pending = deferred<typeof OPEN>(), probe = vi.fn().mockReturnValue(pending.promise);
        const { explorer, render } = create(probe);
        const current = await render(); await idle();
        explorer.setModuleFolder(BOOK, 'M', 'Edited'); explorer.forgetModuleFolder(BOOK, 'M');
        pending.resolve(LOCKED); await flush();
        expect(current).toMatchObject(LOCKED); expect(probe).toHaveBeenCalledTimes(1);
    });
});


describe('protected file icons', () => {
    it('probes a collapsed file without reading modules and draws its closed padlock', async () => {
        const probe = vi.fn().mockResolvedValue({ ...LOCKED, isAccessLocked: true });
        const { explorer } = create(probe);
        const [node] = await explorer.getChildren();
        explorer.getTreeItem(node);
        expect(probe).not.toHaveBeenCalled();
        await idle();
        const item = explorer.getTreeItem(node);
        expect((item.iconPath as { light: { path: string }; dark: { path: string } }).light.path).toContain('/light/file-code-locked.svg');
        expect((item.iconPath as { dark: { path: string } }).dark.path).toContain('/dark/file-code-locked.svg');
        expect(item.tooltip).toMatchObject({ value: expect.stringContaining('VBA project locked') });
    });
    it('shows an open padlock and keeps the signature badge after session authorization', async () => {
        const { explorer, render } = create(async () => ({ ...LOCKED, isAccessLocked: false }));
        const node = await render(); await idle();
        const item = explorer.getTreeItem(node);
        expect((item.iconPath as { light: { path: string } }).light.path).toContain('file-code-unlocked.svg');
        expect(item.description).toContain('[unlocked for this session, signed]');
        expect(item.tooltip).toMatchObject({ value: expect.stringContaining('VBA project unlocked for this session') });
    });
});
