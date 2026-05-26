import * as vscode from 'vscode';
import * as path from 'path';
import { PythonBridge } from './pythonBridge';
import { encodeModuleUri } from './xlideFileSystem';

export type XlideNodeKind = 'xlsm' | 'module' | 'sub';

export interface XlideNode {
    kind: XlideNodeKind;
    label: string;
    /** Absolute path to the .xlsm file. */
    filePath: string;
    /** Module name (for 'module' and 'sub' nodes). */
    moduleName?: string;
    /** Module type: 'standard' | 'class' | 'document' */
    moduleType?: string;
    /** 1-based line number of the procedure (for 'sub' nodes). */
    line?: number;
}

const MODULE_ICONS: Record<string, string> = {
    standard: 'symbol-module',
    class: 'symbol-class',
    document: 'symbol-namespace',
};

export class XlsmExplorer implements vscode.TreeDataProvider<XlideNode> {
    private _emitter = new vscode.EventEmitter<XlideNode | undefined | null | void>();
    readonly onDidChangeTreeData = this._emitter.event;

    constructor(private readonly _bridge: PythonBridge) {}

    refresh(): void {
        this._emitter.fire();
    }

    getTreeItem(node: XlideNode): vscode.TreeItem {
        const item = new vscode.TreeItem(
            node.label,
            node.kind === 'sub'
                ? vscode.TreeItemCollapsibleState.None
                : vscode.TreeItemCollapsibleState.Collapsed,
        );

        switch (node.kind) {
            case 'xlsm':
                item.iconPath = new vscode.ThemeIcon('file-code');
                item.tooltip = node.filePath;
                item.description = path.relative(
                    vscode.workspace.workspaceFolders?.[0]?.uri.fsPath ?? '',
                    path.dirname(node.filePath),
                ) || '';
                item.contextValue = 'xlsm';
                break;

            case 'module':
                item.iconPath = new vscode.ThemeIcon(
                    MODULE_ICONS[node.moduleType ?? 'standard'] ?? 'symbol-module',
                );
                item.description = node.moduleType;
                item.contextValue = 'module';
                item.command = {
                    command: 'xlide.openModule',
                    title: 'Open Module',
                    arguments: [node],
                };
                break;

            case 'sub':
                item.iconPath = new vscode.ThemeIcon('symbol-method');
                item.contextValue = 'sub';
                item.command = {
                    command: 'xlide.openModule',
                    title: 'Go to Procedure',
                    arguments: [node],
                };
                break;
        }

        return item;
    }

    async getChildren(node?: XlideNode): Promise<XlideNode[]> {
        if (!node) {
            return this._getXlsmFiles();
        }
        if (node.kind === 'xlsm') {
            return this._getModules(node.filePath);
        }
        if (node.kind === 'module') {
            return this._getSubs(node.filePath, node.moduleName!);
        }
        return [];
    }

    private async _getXlsmFiles(): Promise<XlideNode[]> {
        const uris = await vscode.workspace.findFiles(
            '**/*.{xlsm,xlsb,xlam}',
            '**/node_modules/**',
        );
        return uris
            .sort((a, b) => a.fsPath.localeCompare(b.fsPath))
            .map((uri) => ({
                kind: 'xlsm' as const,
                label: path.basename(uri.fsPath),
                filePath: uri.fsPath,
            }));
    }

    private async _getModules(filePath: string): Promise<XlideNode[]> {
        try {
            const modules = await this._bridge.call<Array<{ name: string; type: string }>>(
                'listModules',
                { path: filePath },
            );
            return modules.map((m) => ({
                kind: 'module' as const,
                label: m.name,
                filePath,
                moduleName: m.name,
                moduleType: m.type,
            }));
        } catch (err) {
            vscode.window.showErrorMessage(`XLIDE: Failed to list modules in "${path.basename(filePath)}": ${err}`);
            return [];
        }
    }

    private async _getSubs(filePath: string, moduleName: string): Promise<XlideNode[]> {
        try {
            const subs = await this._bridge.call<Array<{ name: string; kind: string; line: number }>>(
                'listSubs',
                { path: filePath, module: moduleName },
            );
            return subs.map((s) => ({
                kind: 'sub' as const,
                label: `${s.kind} ${s.name}`,
                filePath,
                moduleName,
                line: s.line,
            }));
        } catch {
            return [];
        }
    }
}
