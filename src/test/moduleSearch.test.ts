import * as assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { ProjectEngine } from '../projectEngine';
import { XlideFileSystemProvider, encodeModuleUri } from '../xlideFileSystem';
import { ModuleSearchPanel, applyModuleSearchPlan, type PlannedModule } from '../moduleSearchPanel';
import { setExtensionAssetRoot } from '../extensionAssets';
import { readModule, writeModule } from '../vba/projectService';
import { activate, EXTENSION_ID, workspaceRoot, until } from './support';
import type { CommandDeps } from '../commands/shared';

suite('XLIDE Find / Replace panel', () => {
    suiteSetup(async () => { await activate(); });
    for (const [label, extension] of [['Excel', 'xlsm'], ['Word', 'docm'], ['PowerPoint', 'pptm'], ['Access', 'accdb'], ['VB6', 'vbp']]) {
        test(`${label}: real webview worker searches unopened and dirty modules, previews and applies document edits`, async () => {
            const installed = vscode.extensions.getExtension(EXTENSION_ID)!;
            setExtensionAssetRoot(installed.extensionUri.fsPath);
            const file = path.join(workspaceRoot(), `SearchFixture.${extension}`);
            const original = 'Option Explicit\r\nPublic Sub Test()\r\nDebug.Print "XlideNeedle"\r\nEnd Sub\r\n';
            if (extension === 'vbp') {
                fs.writeFileSync(file, 'Type=Exe\r\nModule=SearchUnopened; SearchUnopened.bas\r\nModule=SearchDirty; SearchDirty.bas\r\nName="SearchFixture"\r\n');
                for (const name of ['SearchUnopened', 'SearchDirty']) { fs.writeFileSync(path.join(workspaceRoot(), `${name}.bas`), original); }
            } else {
                fs.copyFileSync(path.join(installed.extensionUri.fsPath, 'assets/templates', `blank.${extension}`), file);
                writeModule(file, 'SearchUnopened', original);
                writeModule(file, 'SearchDirty', original);
            }
            const dirty = await vscode.workspace.openTextDocument(extension === 'vbp'
                ? vscode.Uri.file(path.join(workspaceRoot(), 'SearchDirty.bas'))
                : encodeModuleUri(label === 'Excel' && process.platform === 'win32' ? file.toUpperCase() : file,
                    label === 'Excel' ? 'searchdirty' : 'SearchDirty'));
            const edit = new vscode.WorkspaceEdit();
            edit.insert(dirty.uri, new vscode.Position(dirty.lineCount, 0), "' XlideNeedle unsaved\r\n");
            assert.ok(await vscode.workspace.applyEdit(edit));
            const output = vscode.window.createOutputChannel('Search integration');
            const engine = new ProjectEngine({} as vscode.ExtensionContext, output);
            const provider = new XlideFileSystemProvider(engine);
            const node = { kind: 'project' as const, filePath: file, label };
            const panel = new ModuleSearchPanel({ bridge: engine, fsProvider: provider,
                explorer: { getChildren: async () => [node] } } as unknown as CommandDeps, node);
            const state = panel as unknown as { panel: vscode.WebviewPanel; plan: PlannedModule[];
                send(message: any): void; receive(message: any): Promise<void> };
            const messages: any[] = [];
            const originalSend = state.send.bind(panel);
            state.send = message => { messages.push(message); originalSend(message); };
            try {
                // Drive production controls inside the real webview. All matching
                // still passes through the shipped worker and message handlers.
                const snapshots: any[] = [];
                const snapshotListener = state.panel.webview.onDidReceiveMessage(message => {
                    if (message.type === 'testSnapshotResult') { snapshots.push(message); }
                });
                const html = state.panel.webview.html.replace("send({ type: 'ready' });", `
window.addEventListener('message', event => {
    if (event.origin === window.origin && event.data.type === 'testSnapshot') {
        send({ type: 'testSnapshotResult', pattern: el('pattern').value, replacement: el('replacement').value,
            rows: rows.size, previewRows: document.querySelectorAll('.diff-line').length,
            applyVisible: !el('apply').hidden });
    }
});
send({ type: 'ready' });`);
                const nonce = /script-src 'nonce-([^']+)'/.exec(html)![1];
                state.panel.webview.html = html.replace('</body>', `<script nonce="${nonce}">
window.addEventListener('message', event => {
    if (event.origin === window.origin && event.data.type === 'done') {
        document.getElementById('replacement').value = 'XlideUpdated';
        document.getElementById('all').click();
    }
});
document.getElementById('pattern').value = 'XlideNeedle';
document.getElementById('search').click();
</script></body>`);
                await until(() => {
                    const error = messages.find(m => m.type === 'error' || (m.type === 'done' && m.errors.length));
                    assert.ok(!error, `panel failed: ${JSON.stringify(error)}`);
                    return state.plan.length ? true : undefined;
                }, `real panel should produce a replacement preview (${label})`, 25000);
                assert.equal(state.plan.length, 2);
                assert.equal(state.plan.reduce((n, m) => n + m.edits.length, 0), 3, 'includes the unsaved extra match');
                state.send({ type: 'testSnapshot' });
                await until(() => snapshots.length === 1 ? true : undefined, 'preview DOM should be ready');
                await state.receive({ type: 'open', id: 0 });
                await until(() => !state.panel.visible ? true : undefined, 'following a result should hide the pane');
                state.panel.reveal();
                await until(() => state.panel.visible ? true : undefined, 'search pane should be visible again');
                state.send({ type: 'testSnapshot' });
                await until(() => snapshots.length === 2 ? true : undefined, 'retained webview should respond');
                assert.deepEqual(snapshots[1], snapshots[0], 'query, results and preview survive result navigation');
                assert.equal(snapshots[1].pattern, 'XlideNeedle');
                assert.equal(snapshots[1].rows, 3);
                assert.equal(snapshots[1].previewRows, 6);
                assert.equal(snapshots[1].applyVisible, true);
                snapshotListener.dispose();
                assert.equal(readModule(file, 'SearchUnopened').source, original);
                assert.ok(dirty.getText().includes('unsaved'));
                await applyModuleSearchPlan(state.plan);
                const changed = state.plan.map(module => module.doc);
                assert.ok(changed.every(doc => doc.getText().includes('XlideUpdated') && !doc.getText().includes('XlideNeedle')));
                assert.ok(changed.every(doc => doc.isDirty));
                assert.equal(readModule(file, 'SearchUnopened').source, original, 'preview/apply do not explicitly save the container');
                for (const doc of changed) { assert.ok(await doc.save()); }
                assert.ok(readModule(file, 'SearchUnopened').source.includes('XlideUpdated'));
                assert.ok(readModule(file, 'SearchDirty').source.includes('XlideUpdated unsaved'));
            } finally { panel.dispose(); provider.dispose(); engine.dispose(); output.dispose(); }
        });
    }
});
