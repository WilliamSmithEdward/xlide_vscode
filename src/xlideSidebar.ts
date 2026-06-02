import * as fs from 'fs';
import * as path from 'path';
import * as vscode from 'vscode';
import { resolvedXlideGlobalSettingsFromConfig } from './globalSettings';
import {
    effectiveWorkbookAnalysisSettingsFromConfig,
} from './workbookAnalysisSettings';
import {
    effectiveWorkbookModuleSyncSettingsFromConfig,
} from './workbookModuleSyncSettings';
import {
    isWorkbookSettingsError,
    readWorkbookSettings,
    settingsPathForWorkbook,
} from './workbookSettings';
import { decodeModuleUri, XLIDE_SCHEME } from './xlideFileSystem';
import {
    buildXlideSidebarModel,
    type XlideSidebarActiveWorkbook,
    type XlideSidebarNode,
} from './xlideSidebarModel';

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

        const workbooks = await workbookFiles();
        const activeWorkbook = await activeWorkbookContext(workbooks);
        return buildXlideSidebarModel({
            globalSettings: resolvedXlideGlobalSettingsFromConfig(vscode.workspace.getConfiguration('xlide')),
            hasWorkspace: Boolean(vscode.workspace.workspaceFolders?.length),
            workbookCount: workbooks.length,
            activeWorkbook,
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
        vscode.window.onDidChangeActiveTextEditor(() => provider.refresh()),
        vscode.commands.registerCommand('xlide.openWorkbookSettings', async (settingsPath?: string) => {
            if (!settingsPath) {
                vscode.window.showWarningMessage('XLIDE: No workbook settings file is available.');
                return;
            }
            try {
                const document = await vscode.workspace.openTextDocument(vscode.Uri.file(settingsPath));
                await vscode.window.showTextDocument(document, { preview: false });
            } catch (err) {
                const message = err instanceof Error ? err.message : String(err);
                vscode.window.showErrorMessage(`XLIDE: Could not open workbook settings: ${message}`);
            }
        }),
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
            const watcher = vscode.workspace.createFileSystemWatcher('**/*.xlide_settings.json');
            watcher.onDidCreate(refresh);
            watcher.onDidChange(refresh);
            watcher.onDidDelete(refresh);
            return watcher;
        })(),
    ];
}

async function workbookFileCount(): Promise<number> {
    return (await workbookFiles()).length;
}

async function workbookFiles(): Promise<vscode.Uri[]> {
    const uris = await vscode.workspace.findFiles(
        '**/*.{xlsm,xlsb,xlam}',
        '{**/node_modules/**,**/.venv/**,**/venv/**}',
    );
    return uris
        .filter((uri) => uri.scheme === 'file' && !path.basename(uri.fsPath).startsWith('~$'))
        .sort((left, right) => left.fsPath.localeCompare(right.fsPath));
}

async function activeWorkbookContext(workbooks: readonly vscode.Uri[]): Promise<XlideSidebarActiveWorkbook | undefined> {
    const activeFromEditor = activeWorkbookPathFromEditor();
    if (activeFromEditor) {
        return sidebarWorkbookForPath(activeFromEditor, 'activeEditor');
    }
    if (workbooks.length === 1) {
        return sidebarWorkbookForPath(workbooks[0].fsPath, 'singleWorkbook');
    }
    return undefined;
}

function activeWorkbookPathFromEditor(): string | undefined {
    const editor = vscode.window.activeTextEditor;
    if (!editor) {
        return undefined;
    }
    const uri = editor.document.uri;
    if (uri.scheme !== XLIDE_SCHEME || uri.authority) {
        return undefined;
    }
    return decodeModuleUri(uri).xlsmPath;
}

async function sidebarWorkbookForPath(
    workbookPath: string,
    selectionSource: XlideSidebarActiveWorkbook['selectionSource'],
): Promise<XlideSidebarActiveWorkbook> {
    const settingsPath = settingsPathForWorkbook(workbookPath);
    const base = {
        label: path.basename(workbookPath),
        filePath: workbookPath,
        settingsPath,
        selectionSource,
    };
    try {
        const exists = fs.existsSync(settingsPath);
        const config = await readWorkbookSettings(workbookPath);
        return {
            ...base,
            settingsState: exists ? 'valid' : 'missing',
            analysisSettings: effectiveWorkbookAnalysisSettingsFromConfig(workbookPath, config),
            moduleSyncSettings: effectiveWorkbookModuleSyncSettingsFromConfig(workbookPath, config),
        };
    } catch (err) {
        return {
            ...base,
            settingsState: 'invalid',
            settingsMessage: isWorkbookSettingsError(err)
                ? err.message
                : `Unable to read workbook settings: ${err instanceof Error ? err.message : String(err)}`,
        };
    }
}

export {
    XlideSidebarProvider,
    registerXlideSidebar,
    workbookFileCount,
    workbookFiles,
};
