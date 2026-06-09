import { beforeEach, describe, expect, it, vi } from 'vitest';

const vscodeMock = vi.hoisted(() => ({
    findFiles: vi.fn(),
    showErrorMessage: vi.fn(),
    treeEvents: [] as unknown[],
}));

vi.mock('vscode', () => ({
    EventEmitter: class {
        event = vi.fn();
        fire = vi.fn((node?: unknown) => {
            vscodeMock.treeEvents.push(node);
        });
    },
    MarkdownString: class {
        supportThemeIcons = false;
        constructor(readonly value = '') {}
        appendMarkdown = vi.fn();
    },
    ThemeIcon: class {
        constructor(readonly id: string) {}
    },
    TreeItem: class {
        id?: string;
        iconPath?: unknown;
        tooltip?: unknown;
        description?: string;
        contextValue?: string;
        command?: unknown;
        constructor(readonly label: string, readonly collapsibleState: number) {}
    },
    TreeItemCollapsibleState: {
        None: 0,
        Collapsed: 1,
        Expanded: 2,
    },
    window: {
        showErrorMessage: vscodeMock.showErrorMessage,
    },
    workspace: {
        findFiles: vscodeMock.findFiles,
        workspaceFolders: [{ uri: { fsPath: 'C:\\work' } }],
    },
}));

import { XlsmExplorer } from '../src/xlsmExplorer';

describe('XlsmExplorer', () => {
    beforeEach(() => {
        vscodeMock.findFiles.mockReset();
        vscodeMock.showErrorMessage.mockReset();
        vscodeMock.treeEvents = [];
        vscodeMock.findFiles.mockResolvedValue([
            { scheme: 'file', fsPath: 'C:\\work\\Book.xlsm' },
        ]);
    });

    it('keeps the workbook tree empty until setup is complete', async () => {
        const explorer = new XlsmExplorer(fakeBridge());

        await expect(explorer.getChildren()).resolves.toEqual([]);
        expect(vscodeMock.findFiles).not.toHaveBeenCalled();

        explorer.setSetupComplete(true);
        await expect(explorer.getChildren()).resolves.toMatchObject([{
            kind: 'xlsm',
            label: 'Book.xlsm',
            filePath: 'C:\\work\\Book.xlsm',
        }]);
        expect(vscodeMock.findFiles).toHaveBeenCalledWith(
            '**/*.{xlsm,xlsb,xlam}',
            '{**/node_modules/**,**/.venv/**,**/venv/**}',
        );
    });

    it('coalesces root workbook discovery while the tree is resolving', async () => {
        let resolveFind: (value: Array<{ scheme: string; fsPath: string }>) => void = () => undefined;
        vscodeMock.findFiles.mockReturnValue(new Promise((resolve) => {
            resolveFind = resolve;
        }));
        const explorer = new XlsmExplorer(fakeBridge());
        explorer.setSetupComplete(true);

        const first = explorer.getChildren();
        const second = explorer.getChildren();
        resolveFind([{ scheme: 'file', fsPath: 'C:\\work\\Book.xlsm' }]);

        await expect(Promise.all([first, second])).resolves.toHaveLength(2);
        expect(vscodeMock.findFiles).toHaveBeenCalledTimes(1);
    });

    it('refreshes only affected module nodes during accordion transitions', async () => {
        const explorer = new XlsmExplorer(fakeBridge([
            { name: 'Module1', type: 'standard' },
            { name: 'Module2', type: 'standard' },
        ]));
        explorer.setSetupComplete(true);
        vscodeMock.treeEvents = [];

        const [workbook] = await explorer.getChildren();
        await Promise.resolve();
        vscodeMock.treeEvents = [];
        const [module1, module2] = await explorer.getChildren(workbook);

        explorer.setActiveModule(workbook.filePath, 'Module1');
        expect(vscodeMock.treeEvents).toEqual([module1]);

        vscodeMock.treeEvents = [];
        explorer.setActiveModule(workbook.filePath, 'Module2');
        expect(vscodeMock.treeEvents).toEqual([module1, module2]);
        expect(vscodeMock.treeEvents).not.toContain(workbook);
    });

    it('rotates only the affected module ids when active state changes', async () => {
        const explorer = new XlsmExplorer(fakeBridge([
            { name: 'Module1', type: 'standard' },
            { name: 'Module2', type: 'standard' },
        ]));
        explorer.setSetupComplete(true);
        vscodeMock.treeEvents = [];

        const [workbook] = await explorer.getChildren();
        await Promise.resolve();
        vscodeMock.treeEvents = [];
        const [module1, module2] = await explorer.getChildren(workbook);
        const module1Initial = explorer.getTreeItem(module1);
        const module2Initial = explorer.getTreeItem(module2);

        explorer.setActiveModule(workbook.filePath, 'Module1');
        const module1Active = explorer.getTreeItem(module1);
        const module2Unchanged = explorer.getTreeItem(module2);

        expect(module1Active.id).not.toBe(module1Initial.id);
        expect(module1Active.collapsibleState).toBe(2);
        expect(module2Unchanged.id).toBe(module2Initial.id);
        expect(module2Unchanged.collapsibleState).toBe(1);
    });

    it('collapses loaded non-active workbook roots when the active module changes', async () => {
        vscodeMock.findFiles.mockResolvedValue([
            { scheme: 'file', fsPath: 'C:\\work\\Book1.xlsm' },
            { scheme: 'file', fsPath: 'C:\\work\\Book2.xlsm' },
        ]);
        const explorer = new XlsmExplorer(fakeBridge([
            { name: 'Module1', type: 'standard' },
        ]));
        explorer.setSetupComplete(true);

        const [book1, book2] = await explorer.getChildren();
        const [book1Module] = await explorer.getChildren(book1);
        const [book2Module] = await explorer.getChildren(book2);
        explorer.setActiveModule(book1.filePath, 'Module1');
        const book1Active = explorer.getTreeItem(book1);
        const book2Collapsed = explorer.getTreeItem(book2);

        expect(book1Active.collapsibleState).toBe(2);
        expect(book2Collapsed.collapsibleState).toBe(1);

        vscodeMock.treeEvents = [];
        const book1IdBefore = book1Active.id;
        explorer.setActiveModule(book2.filePath, 'Module1');
        const book1AfterSwitch = explorer.getTreeItem(book1);
        const book2AfterSwitch = explorer.getTreeItem(book2);

        expect(vscodeMock.treeEvents).toContain(book1Module);
        expect(vscodeMock.treeEvents).toContain(book2Module);
        expect(vscodeMock.treeEvents).toContain(book1);
        expect(book1AfterSwitch.id).not.toBe(book1IdBefore);
        expect(book1AfterSwitch.collapsibleState).toBe(1);
        expect(book2AfterSwitch.collapsibleState).toBe(2);
    });

    it('does not re-render other workbook roots when switching modules in the same workbook', async () => {
        vscodeMock.findFiles.mockResolvedValue([
            { scheme: 'file', fsPath: 'C:\\work\\Book1.xlsm' },
            { scheme: 'file', fsPath: 'C:\\work\\Book2.xlsm' },
        ]);
        const explorer = new XlsmExplorer(fakeBridge([
            { name: 'Module1', type: 'standard' },
            { name: 'Module2', type: 'standard' },
        ]));
        explorer.setSetupComplete(true);

        const [book1, book2] = await explorer.getChildren();
        const [book1Module1, book1Module2] = await explorer.getChildren(book1);
        await explorer.getChildren(book2);

        explorer.setActiveModule(book1.filePath, 'Module1');
        vscodeMock.treeEvents = [];
        explorer.setActiveModule(book1.filePath, 'Module2');

        expect(vscodeMock.treeEvents).toEqual([book1Module1, book1Module2]);
        expect(vscodeMock.treeEvents).not.toContain(book2);
    });

    it('does not run protection probes while discovering root workbooks', async () => {
        const call = vi.fn(() => Promise.resolve({ isPasswordProtected: false, isSigned: false }));
        const bridge = { call } as unknown as ConstructorParameters<typeof XlsmExplorer>[0];
        const explorer = new XlsmExplorer(bridge);
        explorer.setSetupComplete(true);

        await explorer.getChildren();
        await explorer.getChildren();

        expect(call).not.toHaveBeenCalled();
    });

    it('coalesces concurrent module list loads for one workbook', async () => {
        let resolveModules: (value: Array<{ name: string; type: string }>) => void = () => undefined;
        const call = vi.fn((method: string) => {
            if (method === 'listModules') {
                return new Promise((resolve) => {
                    resolveModules = resolve;
                });
            }
            return Promise.resolve({ isPasswordProtected: false, isSigned: false });
        });
        const explorer = new XlsmExplorer({ call } as unknown as ConstructorParameters<typeof XlsmExplorer>[0]);
        explorer.setSetupComplete(true);
        const [workbook] = await explorer.getChildren();

        const first = explorer.getChildren(workbook);
        const second = explorer.getChildren(workbook);
        resolveModules([{ name: 'Module1', type: 'standard' }]);

        await expect(Promise.all([first, second])).resolves.toHaveLength(2);
        expect(call.mock.calls.filter(([method]) => method === 'listModules')).toHaveLength(1);
    });

    it('caches procedure lists until a module sub refresh', async () => {
        const bridge = fakeBridge(
            [{ name: 'Module1', type: 'standard' }],
            [{ name: 'Run', kind: 'Sub', line: 1 }],
        );
        const explorer = new XlsmExplorer(bridge);
        explorer.setSetupComplete(true);
        const [workbook] = await explorer.getChildren();
        const [module] = await explorer.getChildren(workbook);

        await explorer.getChildren(module);
        await explorer.getChildren(module);

        expect(vi.mocked(bridge.call).mock.calls.filter(([method]) => method === 'listSubs')).toHaveLength(1);

        explorer.refreshModuleSubs(workbook.filePath, 'Module1');
        await explorer.getChildren(module);

        expect(vi.mocked(bridge.call).mock.calls.filter(([method]) => method === 'listSubs')).toHaveLength(2);
    });

    it('defers protection probes until after tree expansion goes idle', async () => {
        vi.useFakeTimers();
        try {
            const bridge = fakeBridge([{ name: 'Module1', type: 'standard' }]);
            const explorer = new XlsmExplorer(bridge);
            explorer.setSetupComplete(true);
            const [workbook] = await explorer.getChildren();

            await explorer.getChildren(workbook);

            expect(vi.mocked(bridge.call).mock.calls.filter(([method]) => method === 'getProtectionInfo'))
                .toHaveLength(0);

            await vi.advanceTimersByTimeAsync(1999);
            expect(vi.mocked(bridge.call).mock.calls.filter(([method]) => method === 'getProtectionInfo'))
                .toHaveLength(0);

            await vi.advanceTimersByTimeAsync(1);
            expect(vi.mocked(bridge.call).mock.calls.filter(([method]) => method === 'getProtectionInfo'))
                .toHaveLength(1);
        } finally {
            vi.useRealTimers();
        }
    });
});

function fakeBridge(
    modules: Array<{ name: string; type: string }> = [],
    subs: Array<{ name: string; kind: string; line: number }> = [],
) {
    return {
        call: vi.fn((method: string) => {
            if (method === 'listModules') {
                return Promise.resolve(modules);
            }
            if (method === 'listSubs') {
                return Promise.resolve(subs);
            }
            if (method === 'getProtectionInfo') {
                return Promise.resolve({ isPasswordProtected: false, isSigned: false });
            }
            return Promise.resolve({ isPasswordProtected: false, isSigned: false });
        }),
    } as unknown as ConstructorParameters<typeof XlsmExplorer>[0];
}
