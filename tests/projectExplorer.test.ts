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

import { ProjectExplorer } from '../src/projectExplorer';
import { MACRO_CONTAINER_GLOB } from '../src/macroContainerUi';

describe('a VB6 project in the explorer', () => {
    beforeEach(() => {
        vscodeMock.findFiles.mockReset();
        vscodeMock.showErrorMessage.mockReset();
        vscodeMock.treeEvents = [];
        vscodeMock.findFiles.mockResolvedValue([
            { scheme: 'file', fsPath: 'C:\\work\\App.vbp' },
        ]);
    });

    it('lists the .vbp at the root, its modules with their files, and a designer row under a form', async () => {
        const explorer = new ProjectExplorer(fakeBridge(
            [
                { name: 'Form1', type: 'userform', filePath: 'C:\\work\\Form1.frm' },
                { name: 'modMain', type: 'standard', filePath: 'C:\\work\\modMain.bas' },
                { name: 'ctxThing', type: 'usercontrol', filePath: 'C:\\work\\ctxThing.ctl' },
                { name: 'dsrThing', type: 'designer', filePath: 'C:\\work\\dsrThing.dsr' },
            ],
            [{ name: 'Form_Load', kind: 'Sub', line: 12 }],
        ));

        const [project] = await explorer.getChildren();
        expect(project).toMatchObject({ kind: 'project', label: 'App.vbp', filePath: 'C:\\work\\App.vbp' });
        expect(explorer.getTreeItem(project).contextValue).toBe('vb6Project');

        const modules = await explorer.getChildren(project);
        expect(modules.map((m) => [m.moduleName, m.moduleType, m.moduleFilePath])).toEqual([
            ['Form1', 'userform', 'C:\\work\\Form1.frm'],
            ['modMain', 'standard', 'C:\\work\\modMain.bas'],
            ['ctxThing', 'usercontrol', 'C:\\work\\ctxThing.ctl'],
            ['dsrThing', 'designer', 'C:\\work\\dsrThing.dsr'],
        ]);

        // A VB6 form's designer opens over the form's own file; its row sits
        // first, above the handlers, as it does for a UserForm.
        const children = await explorer.getChildren(modules[0]);
        expect(children.map((c) => c.kind)).toEqual(['designer', 'sub']);
        expect(children[0]).toMatchObject({ kind: 'designer', moduleName: 'Form1', filePath: 'C:\\work\\App.vbp' });
        expect(children[1]).toMatchObject({ moduleName: 'Form1', line: 12 });
        // A UserControl has a designer too; an ActiveX Designer (.dsr) does not.
        const controlChildren = await explorer.getChildren(modules[2]);
        expect(controlChildren[0]).toMatchObject({ kind: 'designer', moduleName: 'ctxThing' });
        const designerChildren = await explorer.getChildren(modules[3]);
        expect(designerChildren.map((c) => c.kind)).toEqual(['sub']);
    });
});

describe('ProjectExplorer', () => {
    beforeEach(() => {
        vscodeMock.findFiles.mockReset();
        vscodeMock.showErrorMessage.mockReset();
        vscodeMock.treeEvents = [];
        vscodeMock.findFiles.mockResolvedValue([
            { scheme: 'file', fsPath: 'C:\\work\\Book.xlsm' },
        ]);
    });

    it('lists workspace projects at the tree root', async () => {
        const explorer = new ProjectExplorer(fakeBridge());

        await expect(explorer.getChildren()).resolves.toMatchObject([{
            kind: 'project',
            label: 'Book.xlsm',
            filePath: 'C:\\work\\Book.xlsm',
        }]);
        expect(vscodeMock.findFiles).toHaveBeenCalledWith(
            MACRO_CONTAINER_GLOB,
            '{**/node_modules/**,**/.venv/**,**/venv/**}',
        );
    });

    it('coalesces root project discovery while the tree is resolving', async () => {
        let resolveFind: (value: Array<{ scheme: string; fsPath: string }>) => void = () => undefined;
        vscodeMock.findFiles.mockReturnValue(new Promise((resolve) => {
            resolveFind = resolve;
        }));
        const explorer = new ProjectExplorer(fakeBridge());

        const first = explorer.getChildren();
        const second = explorer.getChildren();
        resolveFind([{ scheme: 'file', fsPath: 'C:\\work\\Book.xlsm' }]);

        await expect(Promise.all([first, second])).resolves.toHaveLength(2);
        expect(vscodeMock.findFiles).toHaveBeenCalledTimes(1);
    });

    it('refreshes only affected module nodes during accordion transitions', async () => {
        const explorer = new ProjectExplorer(fakeBridge([
            { name: 'Module1', type: 'standard' },
            { name: 'Module2', type: 'standard' },
        ]));
        vscodeMock.treeEvents = [];

        const [project] = await explorer.getChildren();
        await Promise.resolve();
        vscodeMock.treeEvents = [];
        const [module1, module2] = await explorer.getChildren(project);

        explorer.setActiveModule(project.filePath, 'Module1');
        expect(vscodeMock.treeEvents).toEqual([module1]);

        vscodeMock.treeEvents = [];
        explorer.setActiveModule(project.filePath, 'Module2');
        expect(vscodeMock.treeEvents).toEqual([module1, module2]);
        expect(vscodeMock.treeEvents).not.toContain(project);
    });

    it('rotates only the affected module ids when active state changes', async () => {
        const explorer = new ProjectExplorer(fakeBridge([
            { name: 'Module1', type: 'standard' },
            { name: 'Module2', type: 'standard' },
        ]));
        vscodeMock.treeEvents = [];

        const [project] = await explorer.getChildren();
        await Promise.resolve();
        vscodeMock.treeEvents = [];
        const [module1, module2] = await explorer.getChildren(project);
        const module1Initial = explorer.getTreeItem(module1);
        const module2Initial = explorer.getTreeItem(module2);

        explorer.setActiveModule(project.filePath, 'Module1');
        const module1Active = explorer.getTreeItem(module1);
        const module2Unchanged = explorer.getTreeItem(module2);

        expect(module1Active.id).not.toBe(module1Initial.id);
        expect(module1Active.collapsibleState).toBe(2);
        expect(module2Unchanged.id).toBe(module2Initial.id);
        expect(module2Unchanged.collapsibleState).toBe(1);
    });

    it('collapses loaded non-active project roots when the active module changes', async () => {
        vscodeMock.findFiles.mockResolvedValue([
            { scheme: 'file', fsPath: 'C:\\work\\Book1.xlsm' },
            { scheme: 'file', fsPath: 'C:\\work\\Book2.xlsm' },
        ]);
        const explorer = new ProjectExplorer(fakeBridge([
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

    it('re-expands a project when focus returns to it (A -> B -> A)', async () => {
        vscodeMock.findFiles.mockResolvedValue([
            { scheme: 'file', fsPath: 'C:\\work\\Book1.xlsm' },
            { scheme: 'file', fsPath: 'C:\\work\\Book2.xlsm' },
        ]);
        const explorer = new ProjectExplorer(fakeBridge([
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
        // Returning to A is a real cross-project switch (previousProjectKey is B,
        // not undefined), so the root refresh fires and A re-expands rather than
        // keeping VS Code's remembered collapsed state.
        expect(vscodeMock.treeEvents).toContain(undefined);
        expect(book1Back.id).not.toBe(book1IdBeforeReturn);
        expect(book1Back.collapsibleState).toBe(2);
        expect(book2Back.collapsibleState).toBe(1);
    });

    it('collapses every other project on a cross-project switch (strict accordion, 3 projects)', async () => {
        vscodeMock.findFiles.mockResolvedValue([
            { scheme: 'file', fsPath: 'C:\\work\\Book1.xlsm' },
            { scheme: 'file', fsPath: 'C:\\work\\Book2.xlsm' },
            { scheme: 'file', fsPath: 'C:\\work\\Book3.xlsm' },
        ]);
        const explorer = new ProjectExplorer(fakeBridge([
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

    it('does not re-render other project roots when switching modules in the same project', async () => {
        vscodeMock.findFiles.mockResolvedValue([
            { scheme: 'file', fsPath: 'C:\\work\\Book1.xlsm' },
            { scheme: 'file', fsPath: 'C:\\work\\Book2.xlsm' },
        ]);
        const explorer = new ProjectExplorer(fakeBridge([
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

    it('does not run protection probes while discovering root projects', async () => {
        const call = vi.fn(() => Promise.resolve({ isPasswordProtected: false, isSigned: false }));
        const bridge = { call } as unknown as ConstructorParameters<typeof ProjectExplorer>[0];
        const explorer = new ProjectExplorer(bridge);

        await explorer.getChildren();
        await explorer.getChildren();

        expect(call).not.toHaveBeenCalled();
    });

    it('coalesces concurrent module list loads for one project', async () => {
        let resolveModules: (value: Array<{ name: string; type: string }>) => void = () => undefined;
        const call = vi.fn((method: string) => {
            if (method === 'listModules') {
                return new Promise((resolve) => {
                    resolveModules = resolve;
                });
            }
            return Promise.resolve({ isPasswordProtected: false, isSigned: false });
        });
        const explorer = new ProjectExplorer({ call } as unknown as ConstructorParameters<typeof ProjectExplorer>[0]);
        const [project] = await explorer.getChildren();

        const first = explorer.getChildren(project);
        const second = explorer.getChildren(project);
        resolveModules([{ name: 'Module1', type: 'standard' }]);

        await expect(Promise.all([first, second])).resolves.toHaveLength(2);
        expect(call.mock.calls.filter(([method]) => method === 'listModules')).toHaveLength(1);
    });

    it('caches procedure lists until a module sub refresh', async () => {
        const bridge = fakeBridge(
            [{ name: 'Module1', type: 'standard' }],
            [{ name: 'Run', kind: 'Sub', line: 1 }],
        );
        const explorer = new ProjectExplorer(bridge);
        const [project] = await explorer.getChildren();
        const [module] = await explorer.getChildren(project);

        await explorer.getChildren(module);
        await explorer.getChildren(module);

        expect(vi.mocked(bridge.call).mock.calls.filter(([method]) => method === 'listSubs')).toHaveLength(1);

        explorer.refreshModuleSubs(project.filePath, 'Module1');
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
        const explorer = new ProjectExplorer({ call } as unknown as ConstructorParameters<typeof ProjectExplorer>[0]);
        const [project] = await explorer.getChildren();
        const [module] = await explorer.getChildren(project);

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
            const explorer = new ProjectExplorer(
                { call } as unknown as ConstructorParameters<typeof ProjectExplorer>[0],
                { appendLine } as unknown as ConstructorParameters<typeof ProjectExplorer>[1],
            );
            const [project] = await explorer.getChildren();

            await explorer.getChildren(project);
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
            const explorer = new ProjectExplorer(bridge);
            const [project] = await explorer.getChildren();

            await explorer.getChildren(project);

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

describe('ProjectExplorer transient load failures', () => {
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
        } as unknown as ConstructorParameters<typeof ProjectExplorer>[0];

        const explorer = new ProjectExplorer(bridge);
        const [project] = await explorer.getChildren();

        // Failure: the node must yield a clickable placeholder, not [] (VS Code
        // caches resolved children, so [] would brick the node until restart).
        const failed = await explorer.getChildren(project);
        expect(failed).toMatchObject([{ kind: 'loadError', filePath: 'C:\\work\\Book.xlsm' }]);
        expect(failed[0].moduleName).toBeUndefined();

        // Excel closed: retryLoad clears the caches and re-fires the workbook
        // node so VS Code re-queries children, which now succeed.
        failNext = false;
        vscodeMock.treeEvents = [];
        explorer.retryLoad(failed[0]);
        expect(vscodeMock.treeEvents).toContainEqual(expect.objectContaining({ kind: 'project' }));
        const recovered = await explorer.getChildren(project);
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
        } as unknown as ConstructorParameters<typeof ProjectExplorer>[0];

        const explorer = new ProjectExplorer(bridge);
        const [project] = await explorer.getChildren();
        const [module] = await explorer.getChildren(project);

        const failed = await explorer.getChildren(module);
        expect(failed).toMatchObject([{ kind: 'loadError', moduleName: 'Module1' }]);

        failNext = false;
        explorer.retryLoad(failed[0]);
        const recovered = await explorer.getChildren(module);
        expect(recovered).toMatchObject([{ kind: 'sub', label: 'Sub DoIt' }]);
    });
});

function fakeBridge(
    modules: Array<{ name: string; type: string; filePath?: string }> = [],
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
    } as unknown as ConstructorParameters<typeof ProjectExplorer>[0];
}
