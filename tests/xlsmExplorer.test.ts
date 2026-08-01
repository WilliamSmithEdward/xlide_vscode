import { beforeEach, describe, expect, it, vi } from 'vitest';

const vscodeMock = vi.hoisted(() => ({
    findFiles: vi.fn(),
    showErrorMessage: vi.fn(),
    treeEvents: [] as unknown[],
}));

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
    EventEmitter: class {
        event = vi.fn();
        fire = vi.fn((node?: unknown) => {
            vscodeMock.treeEvents.push(node);
        });
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

    it('lists workspace workbooks at the tree root', async () => {
        const explorer = new XlsmExplorer(fakeBridge());

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
        // A root refresh (fire(undefined)) is required so VS Code rebuilds the
        // workbook list and actually applies the new collapse/expand states;
        // firing the workbook nodes in place does not reliably collapse them.
        expect(vscodeMock.treeEvents).toContain(undefined);
        expect(book1AfterSwitch.id).not.toBe(book1IdBefore);
        // The now-active workbook's id changes too, so it re-renders Expanded
        // instead of keeping VS Code's remembered collapsed state.
        expect(book2AfterSwitch.id).not.toBe(book2Collapsed.id);
        expect(book1AfterSwitch.collapsibleState).toBe(1);
        expect(book2AfterSwitch.collapsibleState).toBe(2);
    });

    it('re-expands a workbook when focus returns to it (A -> B -> A)', async () => {
        vscodeMock.findFiles.mockResolvedValue([
            { scheme: 'file', fsPath: 'C:\\work\\Book1.xlsm' },
            { scheme: 'file', fsPath: 'C:\\work\\Book2.xlsm' },
        ]);
        const explorer = new XlsmExplorer(fakeBridge([
            { name: 'Module1', type: 'standard' },
        ]));

        const [book1, book2] = await explorer.getChildren();
        await explorer.getChildren(book1);
        await explorer.getChildren(book2);

        explorer.setActiveModule(book1.filePath, 'Module1'); // A active
        explorer.setActiveModule(book2.filePath, 'Module1'); // -> B (A collapses)
        vscodeMock.treeEvents = [];
        const book1IdBeforeReturn = explorer.getTreeItem(book1).id;

        explorer.setActiveModule(book1.filePath, 'Module1'); // -> back to A

        const book1Back = explorer.getTreeItem(book1);
        const book2Back = explorer.getTreeItem(book2);
        // Returning to A is a real cross-workbook switch (previousWorkbookKey is B,
        // not undefined), so the root refresh fires and A re-expands rather than
        // keeping VS Code's remembered collapsed state.
        expect(vscodeMock.treeEvents).toContain(undefined);
        expect(book1Back.id).not.toBe(book1IdBeforeReturn);
        expect(book1Back.collapsibleState).toBe(2);
        expect(book2Back.collapsibleState).toBe(1);
    });

    it('collapses every other workbook on a cross-workbook switch (strict accordion, 3 workbooks)', async () => {
        vscodeMock.findFiles.mockResolvedValue([
            { scheme: 'file', fsPath: 'C:\\work\\Book1.xlsm' },
            { scheme: 'file', fsPath: 'C:\\work\\Book2.xlsm' },
            { scheme: 'file', fsPath: 'C:\\work\\Book3.xlsm' },
        ]);
        const explorer = new XlsmExplorer(fakeBridge([
            { name: 'Module1', type: 'standard' },
        ]));

        const [book1, book2, book3] = await explorer.getChildren();
        await explorer.getChildren(book1);
        await explorer.getChildren(book2);
        await explorer.getChildren(book3);

        explorer.setActiveModule(book1.filePath, 'Module1'); // A active
        vscodeMock.treeEvents = [];
        explorer.setActiveModule(book2.filePath, 'Module1'); // -> B

        // Strict accordion: switching to B collapses BOTH the previous workbook A
        // and the unrelated third workbook C; only the active workbook stays open.
        expect(vscodeMock.treeEvents).toContain(book3); // C is actively re-rendered to collapse
        expect(explorer.getTreeItem(book1).collapsibleState).toBe(1);
        expect(explorer.getTreeItem(book2).collapsibleState).toBe(2);
        expect(explorer.getTreeItem(book3).collapsibleState).toBe(1);
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
        const [workbook] = await explorer.getChildren();
        const [module] = await explorer.getChildren(workbook);

        await explorer.getChildren(module);
        await explorer.getChildren(module);

        expect(vi.mocked(bridge.call).mock.calls.filter(([method]) => method === 'listSubs')).toHaveLength(1);

        explorer.refreshModuleSubs(workbook.filePath, 'Module1');
        await explorer.getChildren(module);

        expect(vi.mocked(bridge.call).mock.calls.filter(([method]) => method === 'listSubs')).toHaveLength(2);
    });

    it('surfaces procedure-list failures like module-list failures', async () => {
        const call = vi.fn((method: string) => {
            if (method === 'listModules') {
                return Promise.resolve([{ name: 'Module1', type: 'standard' }]);
            }
            if (method === 'listSubs') {
                return Promise.reject(new Error('workbook is locked'));
            }
            return Promise.resolve({ isPasswordProtected: false, isSigned: false });
        });
        const explorer = new XlsmExplorer({ call } as unknown as ConstructorParameters<typeof XlsmExplorer>[0]);
        const [workbook] = await explorer.getChildren();
        const [module] = await explorer.getChildren(workbook);

        // Failures yield a retry placeholder (never []): VS Code caches resolved
        // children, so [] would leave the node permanently empty after a
        // transient lock.
        await expect(explorer.getChildren(module)).resolves.toMatchObject([
            { kind: 'loadError', moduleName: 'Module1' },
        ]);

        expect(vscodeMock.showErrorMessage).toHaveBeenCalledWith(
            expect.stringContaining('Failed to list procedures in "Module1"'),
        );
        expect(vscodeMock.showErrorMessage).toHaveBeenCalledWith(
            expect.stringContaining('workbook is locked'),
        );
    });

    it('logs swallowed protection probe failures to the output channel', async () => {
        vi.useFakeTimers();
        try {
            const call = vi.fn((method: string) => {
                if (method === 'listModules') {
                    return Promise.resolve([{ name: 'Module1', type: 'standard' }]);
                }
                if (method === 'getProtectionInfo') {
                    return Promise.reject(new Error('probe failed'));
                }
                return Promise.resolve([]);
            });
            const appendLine = vi.fn();
            const explorer = new XlsmExplorer(
                { call } as unknown as ConstructorParameters<typeof XlsmExplorer>[0],
                { appendLine } as unknown as ConstructorParameters<typeof XlsmExplorer>[1],
            );
            const [workbook] = await explorer.getChildren();

            await explorer.getChildren(workbook);
            await vi.advanceTimersByTimeAsync(2000);

            expect(appendLine).toHaveBeenCalledWith(
                expect.stringContaining('Protection probe failed for "Book.xlsm"'),
            );
            expect(vscodeMock.showErrorMessage).not.toHaveBeenCalled();
        } finally {
            vi.useRealTimers();
        }
    });

    it('defers protection probes until after tree expansion goes idle', async () => {
        vi.useFakeTimers();
        try {
            const bridge = fakeBridge([{ name: 'Module1', type: 'standard' }]);
            const explorer = new XlsmExplorer(bridge);
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

describe('XlsmExplorer transient load failures', () => {
    beforeEach(() => {
        vscodeMock.findFiles.mockReset();
        vscodeMock.showErrorMessage.mockReset();
        vscodeMock.treeEvents = [];
        vscodeMock.findFiles.mockResolvedValue([
            { scheme: 'file', fsPath: 'C:\\work\\Book.xlsm' },
        ]);
    });

    it('returns a retry placeholder (never empty) when listing modules fails, and recovers via retryLoad', async () => {
        // A workbook briefly locked by Excel: the first listModules rejects,
        // the next succeeds once the lock is gone.
        let failNext = true;
        const bridge = {
            call: vi.fn((method: string) => {
                if (method === 'listModules') {
                    if (failNext) {
                        return Promise.reject(new Error('Permission denied: being used by another process'));
                    }
                    return Promise.resolve([{ name: 'Module1', type: 'standard' }]);
                }
                return Promise.resolve({ isPasswordProtected: false, isSigned: false });
            }),
        } as unknown as ConstructorParameters<typeof XlsmExplorer>[0];

        const explorer = new XlsmExplorer(bridge);
        const [workbook] = await explorer.getChildren();

        // Failure: the node must yield a clickable placeholder, not [] (VS Code
        // caches resolved children, so [] would brick the node until restart).
        const failed = await explorer.getChildren(workbook);
        expect(failed).toMatchObject([{ kind: 'loadError', filePath: 'C:\\work\\Book.xlsm' }]);
        expect(failed[0].moduleName).toBeUndefined();

        // Excel closed: retryLoad clears the caches and re-fires the workbook
        // node so VS Code re-queries children, which now succeed.
        failNext = false;
        vscodeMock.treeEvents = [];
        explorer.retryLoad(failed[0]);
        expect(vscodeMock.treeEvents).toContainEqual(expect.objectContaining({ kind: 'xlsm' }));
        const recovered = await explorer.getChildren(workbook);
        expect(recovered).toMatchObject([{ kind: 'module', moduleName: 'Module1' }]);
    });

    it('returns a retry placeholder for a failed sub listing and recovers', async () => {
        let failNext = true;
        const bridge = {
            call: vi.fn((method: string) => {
                if (method === 'listModules') {
                    return Promise.resolve([{ name: 'Module1', type: 'standard' }]);
                }
                if (method === 'listSubs') {
                    if (failNext) {
                        return Promise.reject(new Error('sharing violation'));
                    }
                    return Promise.resolve([{ name: 'DoIt', kind: 'Sub', line: 1 }]);
                }
                return Promise.resolve({ isPasswordProtected: false, isSigned: false });
            }),
        } as unknown as ConstructorParameters<typeof XlsmExplorer>[0];

        const explorer = new XlsmExplorer(bridge);
        const [workbook] = await explorer.getChildren();
        const [module] = await explorer.getChildren(workbook);

        const failed = await explorer.getChildren(module);
        expect(failed).toMatchObject([{ kind: 'loadError', moduleName: 'Module1' }]);

        failNext = false;
        explorer.retryLoad(failed[0]);
        const recovered = await explorer.getChildren(module);
        expect(recovered).toMatchObject([{ kind: 'sub', label: 'Sub DoIt' }]);
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
