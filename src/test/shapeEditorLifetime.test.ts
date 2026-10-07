import * as assert from 'node:assert/strict';
import * as vscode from 'vscode';
import { setExtensionAssetRoot } from '../extensionAssets';
import { openShapeEditor, type ShapeEditorDeps } from '../shapeEditor';
import { closeAllEditors, EXTENSION_ID, until, workbookPath } from './support';

suite('Shape editor opening in the extension host', () => {
    teardown(async () => { await closeAllEditors(); });

    for (const mode of ['existing', 'new'] as const) {
        test(`concurrent ${mode} shape requests create one real webview tab`, async () => {
            const extension = vscode.extensions.getExtension(EXTENSION_ID);
            assert.ok(extension);
            setExtensionAssetRoot(extension.extensionUri.fsPath);
            const output = vscode.window.createOutputChannel('Shape opening regression');
            let release!: () => void;
            const ready = new Promise<void>(yes => { release = yes; });
            let shapeReads = 0, macroReads = 0;
            const shape = { name: 'OpeningProbe', kind: 'shape' as const };
            const deps = {
                bridge: { call: async (method: string) => {
                    if (method === 'listShapes') { shapeReads++; }
                    else if (method === 'shapeMacros') { macroReads++; }
                    await ready;
                    return method === 'listShapes'
                        ? { surfaces: [{ surface: 'Sheet1', shapes: [shape] }] } : { macros: [] };
                } },
                explorer: { refreshShapes: () => {} }, out: output,
                goToMacro: async () => {},
            } as unknown as ShapeEditorDeps;
            const context = { extensionUri: extension.extensionUri } as vscode.ExtensionContext;
            const target = { host: 'excel' as const, surface: 'Sheet1', ...(mode === 'existing' ? { shape } : {}) };
            const tabs = () => vscode.window.tabGroups.all.flatMap(group => group.tabs).filter(tab =>
                tab.input instanceof vscode.TabInputWebview && tab.input.viewType.includes('xlide.shapeEditor'));
            try {
                const requests = Array.from({ length: 8 }, () => openShapeEditor(deps, context, workbookPath(), target));
                release();
                await Promise.all(requests);
                assert.equal(shapeReads, mode === 'existing' ? 1 : 0);
                assert.equal(macroReads, 1);
                await until(() => tabs().length > 0 ? true : undefined, 'shape editor tab should appear');
                assert.equal(tabs().length, 1, 'concurrent requests should share one tab');
            } finally { await closeAllEditors(); output.dispose(); }
        });
    }
});
