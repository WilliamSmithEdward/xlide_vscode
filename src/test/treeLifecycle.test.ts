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
