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
