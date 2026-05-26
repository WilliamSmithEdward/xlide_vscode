import * as vscode from 'vscode';
import * as path from 'path';
import { PythonBridge } from './pythonBridge';
import { XlsmExplorer, XlideNode } from './xlsmExplorer';
import { XlideFileSystemProvider, encodeModuleUri, XLIDE_SCHEME } from './xlideFileSystem';

export function registerCommands(
    _context: vscode.ExtensionContext,
    bridge: PythonBridge,
    explorer: XlsmExplorer,
    _fsProvider: XlideFileSystemProvider,
): vscode.Disposable[] {
    return [
        vscode.commands.registerCommand('xlide.refreshExplorer', () => {
            explorer.refresh();
        }),

        // Open a module (or navigate to a sub's line inside one)
        vscode.commands.registerCommand('xlide.openModule', async (node: XlideNode) => {
            if (!node?.moduleName) { return; }
            const uri = encodeModuleUri(node.filePath, node.moduleName);

            // Set the language to 'vba' so syntax highlighters kick in
            const doc = await vscode.workspace.openTextDocument(uri);
            const editor = await vscode.window.showTextDocument(doc, { preview: false });

            // If a specific line was requested (sub navigation), move cursor there
            if (node.line !== undefined && node.line > 0) {
                const pos = new vscode.Position(node.line - 1, 0);
                editor.selection = new vscode.Selection(pos, pos);
                editor.revealRange(
                    new vscode.Range(pos, pos),
                    vscode.TextEditorRevealType.InCenterIfOutsideViewport,
                );
            }

            // Set language mode to vba for the document
            await vscode.languages.setTextDocumentLanguage(doc, 'vba');
        }),

        // Add a new standard module to an .xlsm
        vscode.commands.registerCommand('xlide.newModule', async (node: XlideNode) => {
            if (node?.kind !== 'xlsm') { return; }
            const name = await vscode.window.showInputBox({
                prompt: 'New module name',
                placeHolder: 'Module1',
                validateInput: (v) =>
                    /^\w+$/.test(v) ? undefined : 'Module names must be alphanumeric',
            });
            if (!name) { return; }

            const stub = `Option Explicit\r\n\r\nSub ${name}_Main()\r\n\r\nEnd Sub\r\n`;
            try {
                await bridge.call('writeModule', {
                    path: node.filePath,
                    module: name,
                    source: stub,
                });
                explorer.refresh();
                // Open the new module immediately
                const uri = encodeModuleUri(node.filePath, name);
                const doc = await vscode.workspace.openTextDocument(uri);
                await vscode.window.showTextDocument(doc, { preview: false });
                await vscode.languages.setTextDocumentLanguage(doc, 'vba');
            } catch (err) {
                vscode.window.showErrorMessage(`XLIDE: Failed to create module: ${err}`);
            }
        }),

        // Rename a module
        vscode.commands.registerCommand('xlide.renameModule', async (node: XlideNode) => {
            if (!node?.moduleName) { return; }
            const newName = await vscode.window.showInputBox({
                prompt: `Rename "${node.moduleName}" to`,
                value: node.moduleName,
                validateInput: (v) =>
                    /^\w+$/.test(v) ? undefined : 'Module names must be alphanumeric',
            });
            if (!newName || newName === node.moduleName) { return; }

            try {
                await bridge.call('renameModule', {
                    path: node.filePath,
                    module: node.moduleName,
                    newName,
                });
                explorer.refresh();
            } catch (err) {
                vscode.window.showErrorMessage(`XLIDE: Rename failed: ${err}`);
            }
        }),

        // Delete a module (with confirmation)
        vscode.commands.registerCommand('xlide.deleteModule', async (node: XlideNode) => {
            if (!node?.moduleName) { return; }
            const choice = await vscode.window.showWarningMessage(
                `Delete module "${node.moduleName}" from "${path.basename(node.filePath)}"?`,
                { modal: true },
                'Delete',
            );
            if (choice !== 'Delete') { return; }

            try {
                await bridge.call('deleteModule', {
                    path: node.filePath,
                    module: node.moduleName,
                });
                // Close any open editors for this module
                const uri = encodeModuleUri(node.filePath, node.moduleName);
                for (const tab of vscode.window.tabGroups.all.flatMap((g) => g.tabs)) {
                    const input = tab.input;
                    if (
                        input instanceof vscode.TabInputText &&
                        input.uri.toString() === uri.toString()
                    ) {
                        await vscode.window.tabGroups.close(tab);
                    }
                }
                explorer.refresh();
            } catch (err) {
                vscode.window.showErrorMessage(`XLIDE: Delete failed: ${err}`);
            }
        }),
    ];
}
