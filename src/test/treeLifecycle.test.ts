import * as assert from 'node:assert/strict';
import { ProjectExplorer } from '../projectExplorer';
import { until, workbookPath } from './support';

function deferred<T>() {
    let resolve!: (value: T) => void;
    const promise = new Promise<T>(yes => { resolve = yes; });
    return { promise, resolve };
}

suite('Explorer lifetime in the extension host', () => {
    test('a late sheet catalog does not relabel disposed module rows', async () => {
        const catalog = deferred<{ sheets: Array<{ name: string; codeName: string; kind: string }> }>();
        let started = false;
        const explorer = new ProjectExplorer({ call: (method: string) => {
            if (method === 'listWorkbookSheets') { started = true; return catalog.promise; }
            if (method === 'listModules') { return Promise.resolve([{ name: 'Sheet1', type: 'document' }]); }
            return Promise.resolve([]);
        } } as unknown as ConstructorParameters<typeof ProjectExplorer>[0]);
        try {
            const project = (await explorer.getChildren()).find(node => node.filePath === workbookPath());
            assert.ok(project);
            const loading = explorer.getChildren(project);
            await until(() => started ? true : undefined, 'catalog read should start');
            const module = explorer.getModuleNode(workbookPath(), 'Sheet1');
            assert.ok(module);
            explorer.dispose();
            catalog.resolve({ sheets: [{ name: 'Visible', codeName: 'Sheet1', kind: 'worksheet' }] });
            await loading;
            assert.equal(module.label, 'Sheet1');
            assert.equal(module.sheetName, undefined);
        } finally {
            explorer.dispose();
        }
    });

    test('a late module read starts no new work after disposal', async () => {
        const modules = deferred<Array<{ name: string; type: string }>>();
        const calls: string[] = [];
        const explorer = new ProjectExplorer({ call: (method: string) => {
            calls.push(method);
            return method === 'listModules' ? modules.promise : Promise.resolve([]);
        } } as unknown as ConstructorParameters<typeof ProjectExplorer>[0]);
        try {
            const project = (await explorer.getChildren()).find(node => node.filePath === workbookPath());
            assert.ok(project, 'fixture workbook should be discoverable');
            const loading = explorer.getChildren(project);
            await until(() => calls.includes('listModules') ? true : undefined, 'module read should start');
            explorer.dispose();
            const before = [...calls];
            modules.resolve([{ name: 'Slow', type: 'standard' }]);
            assert.deepEqual(await loading, [], 'disposed providers should not return new rows');
            assert.deepEqual(calls, before, 'a late read must not start sheet or protection reads');
            assert.equal(explorer.getModuleNode(workbookPath(), 'Slow'), undefined);
        } finally {
            explorer.dispose();
        }
    });

    test('a late protection read does not mutate disposed rows', async () => {
        const protection = deferred<{ isPasswordProtected: boolean; isSigned: boolean }>();
        let started = false;
        const explorer = new ProjectExplorer({ call: (method: string) => {
            if (method === 'getProtectionInfo') { started = true; return protection.promise; }
            if (method === 'listModules') { return Promise.resolve([{ name: 'M', type: 'standard' }]); }
            return Promise.resolve([]);
        } } as unknown as ConstructorParameters<typeof ProjectExplorer>[0]);
        try {
            const project = (await explorer.getChildren()).find(node => node.filePath === workbookPath());
            assert.ok(project);
            await explorer.getChildren(project);
            await until(() => started ? true : undefined, 'protection read should start');
            explorer.dispose();
            protection.resolve({ isPasswordProtected: true, isSigned: true });
            await new Promise(resolve => setTimeout(resolve, 50));
            assert.equal(project.isPasswordProtected, undefined);
            assert.equal(project.isSigned, undefined);
        } finally {
            explorer.dispose();
        }
    });
});

suite('Explorer shape refresh in the extension host', () => {
    test('retained nested groups refresh without waiting for their parent folder', async () => {
        let name = 'Old', removed = false, reads = 0;
        const explorer = new ProjectExplorer({ call: () => {
            reads++;
            return Promise.resolve({ surfaces: [{ surface: 'Data', shapes: removed ? [] : [
                { name: 'Pair', kind: 'group', shapes: [
                    { name: 'Nested', kind: 'group', shapes: [{ name, kind: 'shape' }] },
                ] },
            ] }] });
        } } as unknown as ConstructorParameters<typeof ProjectExplorer>[0]);
        try {
            const folder = { kind: 'shapes' as const, shapeFolder: 'surface' as const,
                surface: 'Data', label: 'Shapes', filePath: workbookPath() };
            const [group] = await explorer.getChildren(folder);
            const [nested] = await explorer.getChildren(group);
            assert.deepEqual((await explorer.getChildren(nested)).map(row => row.label), ['Old']);
            name = 'New';
            explorer.refreshShapes(workbookPath());
            const [member] = await explorer.getChildren(nested);
            assert.equal(member.label, 'New');
            assert.equal(explorer.getParent(member), nested);
            assert.deepEqual(member.shapePath, ['Pair', 'Nested', 'New']);
            for (let i = 0; i < 100; i++) { await explorer.getChildren(nested); }
            assert.equal(reads, 2, 'warm group expansion must reuse the snapshot');
            removed = true;
            explorer.refreshShapes(workbookPath());
            assert.deepEqual(await explorer.getChildren(group), []);
            assert.deepEqual(await explorer.getChildren(nested), []);
            assert.equal(reads, 3, 'retained groups must share the current listing');
        } finally { explorer.dispose(); }
    });

    test('a shape refresh visits only opened rows in its own project', async () => {
        let shapeCalls = 0, unrelatedPathReads = 0;
        const explorer = new ProjectExplorer({ call: () => {
            shapeCalls++;
            return Promise.resolve({ surfaces: [] });
        } } as unknown as ConstructorParameters<typeof ProjectExplorer>[0]);
        const target = { kind: 'shapes' as const, shapeFolder: 'surface' as const,
            surface: 'Data', label: 'Shapes', filePath: workbookPath() };
        const fired: unknown[] = [];
        const subscription = explorer.onDidChangeTreeData(node => fired.push(node));
        try {
            for (let i = 0; i < 1000; i++) {
                const other = { ...target, surface: `Sheet ${i}` };
                Object.defineProperty(other, 'filePath', { get: () => {
                    unrelatedPathReads++;
                    return `${workbookPath()}.other.xlsm`;
                } });
                await explorer.getChildren(other);
            }
            await explorer.getChildren(target);
            await explorer.getChildren(target);
            unrelatedPathReads = 0;
            for (let i = 0; i < 10; i++) { explorer.refreshShapes(workbookPath()); }
            assert.equal(unrelatedPathReads, 0, 'refresh must not scan other projects');
            assert.deepEqual(fired, Array(10).fill(target));
            assert.equal(shapeCalls, 2, 'refresh notification itself should start no bridge reads');
        } finally { subscription.dispose(); explorer.dispose(); }
    });

    test('a retained bare-sheet folder refreshes before Sheets redraws', async () => {
        let name = 'Data', hasShapes = false;
        const explorer = new ProjectExplorer({ call: (method: string) => {
            if (method === 'listModules') { return Promise.resolve([]); }
            if (method === 'listWorkbookSheets') { return Promise.resolve({ sheets: [{ name, kind: 'worksheet' }] }); }
            if (method === 'listShapes') { return Promise.resolve({ surfaces: hasShapes
                ? [{ surface: name, shapes: [{ name: 'Box', kind: 'shape' }] }] : [] }); }
            return Promise.resolve([]);
        } } as unknown as ConstructorParameters<typeof ProjectExplorer>[0]);
        try {
            const project = (await explorer.getChildren()).find(node => node.filePath === workbookPath());
            assert.ok(project);
            const [sheets] = await explorer.getChildren(project);
            const [bare] = await explorer.getChildren(sheets);
            assert.equal(bare.shapeFolder, 'bareSheets');
            assert.deepEqual((await explorer.getChildren(bare)).map(node => node.label), ['Data']);
            name = 'Renamed';
            explorer.refreshShapes(workbookPath());
            assert.deepEqual((await explorer.getChildren(bare)).map(node => node.label), ['Renamed']);
            hasShapes = true;
            explorer.refreshShapes(workbookPath(), { shapesChanged: true });
            assert.deepEqual(await explorer.getChildren(bare), [], 'the sheet has moved directly under Sheets');
        } finally { explorer.dispose(); }
    });

    test('an overtaken sheet read cannot undo a newer rename', async () => {
        const old = deferred<{ sheets: Array<{ name: string; codeName: string; kind: string }> }>();
        let catalogCalls = 0;
        const explorer = new ProjectExplorer({ call: (method: string) => {
            if (method === 'listModules') { return Promise.resolve([{ name: 'Sheet1', type: 'document' }]); }
            if (method === 'listWorkbookSheets') {
                return ++catalogCalls === 1 ? old.promise
                    : Promise.resolve({ sheets: [{ name: 'New', codeName: 'Sheet1', kind: 'worksheet' }] });
            }
            return Promise.resolve([]);
        } } as unknown as ConstructorParameters<typeof ProjectExplorer>[0]);
        try {
            const project = (await explorer.getChildren()).find(node => node.filePath === workbookPath());
            assert.ok(project);
            const pending = explorer.getChildren(project);
            await until(() => catalogCalls > 0 ? true : undefined, 'catalog read should start');
            explorer.refreshShapes(workbookPath());
            await explorer.getChildren(project);
            old.resolve({ sheets: [{ name: 'Old', codeName: 'Sheet1', kind: 'worksheet' }] });
            await pending;
            assert.equal(explorer.getModuleNode(workbookPath(), 'Sheet1')?.label, 'Sheet1 (New)');
            assert.equal(catalogCalls, 2, 'the old render should reuse the current catalog');
        } finally { explorer.dispose(); }
    });

    test('an overtaken shape read returns the current rows', async () => {
        const old = deferred<{ surfaces: Array<{ surface: string; shapes: Array<{ name: string; kind: string }> }> }>();
        let shapeCalls = 0;
        const explorer = new ProjectExplorer({ call: () => ++shapeCalls === 1 ? old.promise
            : Promise.resolve({ surfaces: [{ surface: 'Data', shapes: [{ name: 'New', kind: 'shape' }] }] })
        } as unknown as ConstructorParameters<typeof ProjectExplorer>[0]);
        try {
            const folder = { kind: 'shapes' as const, shapeFolder: 'surface' as const,
                surface: 'Data', label: 'Shapes', filePath: workbookPath() };
            const pending = explorer.getChildren(folder);
            await until(() => shapeCalls > 0 ? true : undefined, 'shape read should start');
            explorer.refreshShapes(workbookPath());
            assert.deepEqual((await explorer.getChildren(folder)).map(node => node.label), ['New']);
            old.resolve({ surfaces: [{ surface: 'Data', shapes: [{ name: 'Old', kind: 'shape' }] }] });
            assert.deepEqual((await pending).map(node => node.label), ['New']);
            assert.equal(shapeCalls, 2);
        } finally { explorer.dispose(); }
    });
});

suite('Explorer sheet context in the extension host', () => {
    test('an opened Shapes folder targets the renamed sheet', async () => {
        let name = 'Data';
        const explorer = new ProjectExplorer({ call: (method: string) => {
            if (method === 'listModules') { return Promise.resolve([{ name: 'Sheet1', type: 'document' }]); }
            if (method === 'listWorkbookSheets') { return Promise.resolve({ sheets: [{ name, codeName: 'Sheet1', kind: 'worksheet' }] }); }
            if (method === 'listShapes') { return Promise.resolve({ surfaces: [{ surface: name, codeName: 'Sheet1', shapes: [{ name: 'Box', kind: 'shape' }] }] }); }
            return Promise.resolve([]);
        } } as unknown as ConstructorParameters<typeof ProjectExplorer>[0]);
        try {
            const project = (await explorer.getChildren()).find(node => node.filePath === workbookPath());
            assert.ok(project);
            await explorer.getChildren(project);
            const module = explorer.getModuleNode(workbookPath(), 'Sheet1');
            assert.ok(module);
            const [folder] = await explorer.getChildren(module);
            assert.equal(folder.kind, 'shapes');
            await explorer.getChildren(folder);
            name = 'Renamed';
            explorer.refreshShapes(workbookPath());
            assert.deepEqual(await explorer.shapeSurfaceOf(folder), { host: 'excel', surface: 'Renamed' });
        } finally { explorer.dispose(); }
    });

    test('a flat fallback module has its actual project parent', async () => {
        let busy = false;
        const explorer = new ProjectExplorer({ call: (method: string) => {
            if (method === 'listModules') { return Promise.resolve([{ name: 'Sheet1', type: 'document' }]); }
            if (method === 'listWorkbookSheets') { return busy ? Promise.reject(new Error('Workbook busy'))
                : Promise.resolve({ sheets: [{ name: 'Data', codeName: 'Sheet1', kind: 'worksheet' }] }); }
            return Promise.resolve([]);
        } } as unknown as ConstructorParameters<typeof ProjectExplorer>[0]);
        try {
            const project = (await explorer.getChildren()).find(node => node.filePath === workbookPath());
            assert.ok(project);
            await explorer.getChildren(project);
            const module = explorer.getModuleNode(workbookPath(), 'Sheet1');
            assert.ok(module);
            assert.equal(explorer.getParent(module)?.shapeFolder, 'sheets');
            busy = true;
            explorer.refreshShapes(workbookPath());
            assert.deepEqual(await explorer.getChildren(project), [module]);
            assert.equal(explorer.getParent(module), project);
            assert.equal(module.label, 'Sheet1');
        } finally { explorer.dispose(); }
    });
});

suite('Explorer module resolution in the extension host', () => {
    test('a retained project row uses the current root after refresh', async () => {
        const explorer = new ProjectExplorer({ call: (method: string) => method === 'listModules'
            ? Promise.resolve([{ name: 'Sheet1', type: 'document' }])
            : Promise.resolve({ sheets: [{ name: 'Data', codeName: 'Sheet1', kind: 'worksheet' }] })
        } as unknown as ConstructorParameters<typeof ProjectExplorer>[0]);
        try {
            const oldProject = (await explorer.getChildren()).find(node => node.filePath === workbookPath());
            assert.ok(oldProject);
            await explorer.getChildren(oldProject);
            explorer.refresh();
            const currentProject = (await explorer.getChildren()).find(node => node.filePath === workbookPath());
            assert.ok(currentProject);
            await explorer.getChildren(oldProject);
            const module = await explorer.resolveModuleNode(workbookPath(), 'Sheet1');
            assert.ok(module);
            const sheets = explorer.getParent(module);
            assert.ok(sheets);
            assert.equal(explorer.getParent(sheets), currentProject);
        } finally { explorer.dispose(); }
    });

    test('follow waits until the sheet parent has been constructed', async () => {
        const catalog = deferred<{ sheets: Array<{ name: string; codeName: string; kind: string }> }>();
        let started = false;
        const explorer = new ProjectExplorer({ call: (method: string) => {
            if (method === 'listModules') { return Promise.resolve([{ name: 'Sheet1', type: 'document' }]); }
            if (method === 'listWorkbookSheets') { started = true; return catalog.promise; }
            return Promise.resolve([]);
        } } as unknown as ConstructorParameters<typeof ProjectExplorer>[0]);
        try {
            const project = (await explorer.getChildren()).find(node => node.filePath === workbookPath());
            assert.ok(project);
            const drawing = explorer.getChildren(project);
            await until(() => started ? true : undefined, 'sheet layout should be pending');
            let settled = false;
            const follow = explorer.resolveModuleNode(workbookPath(), 'Sheet1').then(node => { settled = true; return node; });
            await new Promise(resolve => setTimeout(resolve, 50));
            const premature = settled;
            catalog.resolve({ sheets: [{ name: 'Data', codeName: 'Sheet1', kind: 'worksheet' }] });
            await drawing;
            const module = await follow;
            assert.equal(premature, false, 'follow must wait for the reveal path');
            assert.ok(module);
            assert.equal(explorer.getParent(module)?.shapeFolder, 'sheets');
        } finally { explorer.dispose(); }
    });

    test('follow rebuilds the parent path when an editor folder override is forgotten', async () => {
        const explorer = new ProjectExplorer({ call: (method: string) => method === 'listModules'
            ? Promise.resolve([{ name: 'M', type: 'standard', folder: 'Saved' }]) : Promise.resolve({ sheets: [] })
        } as unknown as ConstructorParameters<typeof ProjectExplorer>[0]);
        try {
            explorer.setView('folders');
            const module = await explorer.resolveModuleNode(workbookPath(), 'M');
            assert.ok(module);
            explorer.setModuleFolder(workbookPath(), 'M', 'Edited');
            explorer.forgetModuleFolder(workbookPath(), 'M');
            assert.equal(await explorer.resolveModuleNode(workbookPath(), 'M'), module);
            assert.equal(explorer.getParent(module)?.folder, 'Saved');
        } finally { explorer.dispose(); }
    });
});
