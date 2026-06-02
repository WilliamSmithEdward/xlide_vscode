import * as path from 'path';
import * as vscode from 'vscode';
import { resolvedXlideGlobalSettingsFromConfig } from './globalSettings';
import { buildXlideSidebarModel, type XlideSidebarNode } from './xlideSidebarModel';

interface XlideSidebarTreeNode extends XlideSidebarNode {
    children?: XlideSidebarTreeNode[];
}

class XlideSidebarProvider implements vscode.TreeDataProvider<XlideSidebarTreeNode> {
    private readonly _emitter = new vscode.EventEmitter<XlideSidebarTreeNode | undefined | null | void>();
    readonly onDidChangeTreeData = this._emitter.event;

    refresh(): void {
        this._emitter.fire();
    }

    getTreeItem(node: XlideSidebarTreeNode): vscode.TreeItem {
        const item = new vscode.TreeItem(
            node.label,
            node.children && node.children.length > 0
                ? vscode.TreeItemCollapsibleState.Expanded
                : vscode.TreeItemCollapsibleState.None,
        );
        item.id = node.id;
        item.description = node.description;
        item.tooltip = node.tooltip;
        item.contextValue = `xlide-sidebar-${node.kind}`;
        item.iconPath = node.icon ? new vscode.ThemeIcon(node.icon) : undefined;
        item.command = node.command;
        return item;
    }

    async getChildren(node?: XlideSidebarTreeNode): Promise<XlideSidebarTreeNode[]> {
        if (node) {
            return node.children ?? [];
        }

        const workbookCount = await workbookFileCount();
        return buildXlideSidebarModel({
            globalSettings: resolvedXlideGlobalSettingsFromConfig(vscode.workspace.getConfiguration('xlide')),
            hasWorkspace: Boolean(vscode.workspace.workspaceFolders?.length),
            workbookCount,
        }) as XlideSidebarTreeNode[];
    }
}

function registerXlideSidebar(): vscode.Disposable[] {
    const provider = new XlideSidebarProvider();
    const tree = vscode.window.createTreeView('xlide.sidebar', {
        treeDataProvider: provider,
        showCollapseAll: true,
    });

    return [
        tree,
        vscode.workspace.onDidChangeConfiguration((event) => {
            if (event.affectsConfiguration('xlide')) {
                provider.refresh();
            }
        }),
        vscode.workspace.onDidChangeWorkspaceFolders(() => provider.refresh()),
        (() => {
            let timer: ReturnType<typeof setTimeout> | undefined;
            const refresh = () => {
                if (timer !== undefined) {
                    clearTimeout(timer);
                }
                timer = setTimeout(() => {
                    timer = undefined;
                    provider.refresh();
                }, 200);
            };
            const watcher = vscode.workspace.createFileSystemWatcher('**/*.{xlsm,xlsb,xlam}');
            watcher.onDidCreate(refresh);
            watcher.onDidDelete(refresh);
            return watcher;
        })(),
    ];
}

async function workbookFileCount(): Promise<number> {
    const uris = await vscode.workspace.findFiles(
        '**/*.{xlsm,xlsb,xlam}',
        '{**/node_modules/**,**/.venv/**,**/venv/**}',
    );
    return uris.filter((uri) =>
        uri.scheme === 'file' && !path.basename(uri.fsPath).startsWith('~$'),
    ).length;
}

export {
    XlideSidebarProvider,
    registerXlideSidebar,
    workbookFileCount,
};
