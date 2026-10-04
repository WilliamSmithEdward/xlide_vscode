import * as assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { setExtensionAssetRoot } from '../extensionAssets';
import { ProjectEngine } from '../projectEngine';
import { openShapeEditor, type ShapeEditorDeps } from '../shapeEditor';
import { editShape, listShapes } from '../vba/projectService';
import { closeAllEditors, EXTENSION_ID, until, workspaceRoot } from './support';

suite('Shape editor target identity in the extension host', () => {
    teardown(async () => { await closeAllEditors(); });

    for (const first of ['existing', 'new'] as const) {
        test(`Add Shape and a shape named + have separate tabs when opening ${first} first`, async () => {
            const extension = vscode.extensions.getExtension(EXTENSION_ID);
            assert.ok(extension);
            setExtensionAssetRoot(extension.extensionUri.fsPath);
            const file = path.join(workspaceRoot(), 'EditorIdentityFixture.pptm');
            fs.copyFileSync(path.join(extension.extensionUri.fsPath, 'tests/fixtures/binaries/PowerPointShapesFixture.pptm'), file);
            editShape(file, 'Slide 1', { action: 'update', name: 'Badge', newName: '+' });
            const shape = listShapes(file, 'Slide 1').surfaces[0].shapes.find(candidate => candidate.name === '+');
            assert.ok(shape);
            const output = vscode.window.createOutputChannel('Shape identity regression');
            const deps = {
                bridge: new ProjectEngine({} as never), explorer: { refreshShapes: () => {} },
                out: output, goToMacro: async () => {},
            } as unknown as ShapeEditorDeps;
            const context = { extensionUri: extension.extensionUri } as vscode.ExtensionContext;
            const existing = { host: 'powerpoint' as const, surface: 'Slide 1', shape };
            const adding = { host: 'powerpoint' as const, surface: 'Slide 1' };
            const tabs = () => vscode.window.tabGroups.all.flatMap(group => group.tabs).filter(tab =>
                tab.input instanceof vscode.TabInputWebview && tab.input.viewType.includes('xlide.shapeEditor'));
            try {
                await openShapeEditor(deps, context, file, first === 'existing' ? existing : adding);
                await openShapeEditor(deps, context, file, first === 'existing' ? adding : existing);
                await until(() => tabs().length === 2 ? true : undefined, 'both editor targets should have their own tab');
                assert.deepEqual(tabs().map(tab => tab.label).sort(), ['+ - Shape', 'New Shape - Slide 1'].sort());
                await openShapeEditor(deps, context, file, existing);
                await openShapeEditor(deps, context, file, adding);
                assert.equal(tabs().length, 2, 'each target should reuse its own tab');
            } finally { await closeAllEditors(); output.dispose(); fs.unlinkSync(file); }
        });
    }
});
