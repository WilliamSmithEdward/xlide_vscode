import * as vscode from 'vscode';
import { XlsmExplorer } from './xlsmExplorer';
import { XlideFileSystemProvider, XLIDE_SCHEME } from './xlideFileSystem';
import { PythonBridge } from './pythonBridge';
import { registerAgentTools } from './agentTools';
import { registerCommands } from './commands';

export function activate(context: vscode.ExtensionContext): void {
    const out = vscode.window.createOutputChannel('XLIDE');
    out.appendLine('XLIDE activating...');

    const bridge = new PythonBridge(context, out);
    const fsProvider = new XlideFileSystemProvider(bridge);
    const explorer = new XlsmExplorer(bridge);

    context.subscriptions.push(
        out,

        // Virtual read/write filesystem for xlide-vba:// URIs
        vscode.workspace.registerFileSystemProvider(XLIDE_SCHEME, fsProvider, {
            isCaseSensitive: process.platform !== 'win32',
            isReadonly: false,
        }),

        // Tree view in the Explorer sidebar
        vscode.window.createTreeView('xlide.explorer', {
            treeDataProvider: explorer,
            showCollapseAll: true,
        }),

        ...registerCommands(context, bridge, explorer, fsProvider),
        ...registerAgentTools(context, bridge, explorer),

        bridge,
    );

    bridge.start().then(() => {
        out.appendLine('XLIDE ready.');
    }).catch((err: Error) => {
        out.appendLine(`ERROR: Python backend failed to start - ${err.message}`);
        vscode.window.showErrorMessage(
            `XLIDE: Failed to start Python backend - ${err.message}. ` +
            `Check the xlide.pythonPath setting or view the XLIDE output channel.`,
        );
    });
}

export function deactivate(): void { /* nothing async needed */ }
