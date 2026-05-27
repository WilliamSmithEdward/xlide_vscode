import * as vscode from 'vscode';
import * as path from 'path';
import { PythonBridge } from './pythonBridge';
import { encodeModuleUri } from './xlideFileSystem';
import type { LiveShareIntegration } from './liveShare';

export type XlideNodeKind = 'xlsm' | 'module' | 'sub';

export interface XlideNode {
    kind: XlideNodeKind;
    label: string;
    /** Absolute path to the .xlsm file (local) or '' for remote (Live Share guest) nodes. */
    filePath: string;
    /** Module name (for 'module' and 'sub' nodes). */
    moduleName?: string;
    /** Module type: 'standard' | 'class' | 'document' */
    moduleType?: string;
    /** 1-based line number of the procedure (for 'sub' nodes). */
    line?: number;
    /** True when this node refers to a workbook hosted on a Live Share peer. */
    isRemote?: boolean;
    /** Stable id of the remote workbook (only when isRemote). */
    remoteId?: string;
    /** Relative folder for display (remote only). */
    remoteRelativeFolder?: string;
}

const MODULE_ICONS: Record<string, string> = {
    standard: 'symbol-module',
    class: 'symbol-class',
    document: 'symbol-namespace',
    userform: 'window',
};

export class XlsmExplorer implements vscode.TreeDataProvider<XlideNode> {
    private _emitter = new vscode.EventEmitter<XlideNode | undefined | null | void>();
    readonly onDidChangeTreeData = this._emitter.event;

    private _liveShare: LiveShareIntegration | undefined;

    // Stable node references required by treeView.reveal()
    private _xlsmNodes = new Map<string, XlideNode>(); // key: filePath
    private _moduleNodes = new Map<string, XlideNode>(); // key: filePath + '::' + moduleName
    // listModules cache: avoids repeated bridge round-trips while the tree is
    // expanded.  Cleared on refresh() so edits always re-fetch.
    private _modulesListCache = new Map<string, Array<{ name: string; type: string }>>();
    // Accordion: only one module node is expanded at a time.
    private _activeModuleKey: string | undefined;

    constructor(private readonly _bridge: PythonBridge) {}

    setLiveShare(liveShare: LiveShareIntegration): void {
        this._liveShare = liveShare;
        liveShare.onDidChange(() => this.refresh());
    }

    refresh(): void {
        this._xlsmNodes.clear();
        this._moduleNodes.clear();
        this._modulesListCache.clear();
        this._emitter.fire();
    }

    /** Required by treeView.reveal() — walks xlsm -> module -> sub. */
    getParent(node: XlideNode): XlideNode | undefined {
        if (node.kind === 'module') {
            return this._xlsmNodes.get(node.filePath);
        }
        if (node.kind === 'sub') {
            return this._moduleNodes.get(`${node.filePath}::${node.moduleName ?? ''}`);
        }
        return undefined;
    }

    /** Returns the cached module node, if the tree has loaded it. */
    getModuleNode(filePath: string, moduleName: string): XlideNode | undefined {
        return this._moduleNodes.get(`${filePath}::${moduleName}`);
    }

    /** Returns the cached xlsm node, if the tree has loaded it. */
    getXlsmNode(filePath: string): XlideNode | undefined {
        return this._xlsmNodes.get(filePath);
    }

    /**
     * Accordion-expand the given module and collapse all sibling module nodes
     * under the same workbook. Safe to call before the tree has loaded.
     */
    setActiveModule(filePath: string, moduleName: string): void {
        const key = `${filePath}::${moduleName}`;
        if (this._activeModuleKey === key) { return; }
        this._activeModuleKey = key;
        // VS Code applies a TreeItem's collapsibleState only when it first sees
        // an element id.  getTreeItem() therefore stamps an id that encodes the
        // active-state — toggling the active module changes every sibling's id,
        // forcing VS Code to render them fresh with Collapsed.
        const xlsmNode = this._xlsmNodes.get(filePath);
        if (xlsmNode) {
            this._emitter.fire(xlsmNode);
        }
    }

    /**
     * Clear the active module (e.g. when the last XLIDE editor closes) so that
     * every module under every workbook collapses.
     */
    clearActiveModule(): void {
        if (this._activeModuleKey === undefined) { return; }
        this._activeModuleKey = undefined;
        // Re-render every loaded workbook so all module ids rotate and apply
        // the Collapsed initial state.
        for (const xlsmNode of this._xlsmNodes.values()) {
            this._emitter.fire(xlsmNode);
        }
    }

    /**
     * Eagerly loads and caches the root xlsm nodes without waiting for the tree
     * to expand them. Returns the first node (if any) so callers can auto-reveal.
     */
    async warmXlsmCache(): Promise<XlideNode | undefined> {
        const nodes = await this._getXlsmFiles();
        return nodes[0];
    }

    getTreeItem(node: XlideNode): vscode.TreeItem {
        const isActiveModule =
            node.kind === 'module' &&
            `${node.filePath}::${node.moduleName}` === this._activeModuleKey;
        const item = new vscode.TreeItem(
            node.label,
            node.kind === 'sub'
                ? vscode.TreeItemCollapsibleState.None
                : isActiveModule
                    ? vscode.TreeItemCollapsibleState.Expanded
                    : vscode.TreeItemCollapsibleState.Collapsed,
        );

        // Encode the *current* active-module key into every module id.  Each
        // accordion transition therefore changes every sibling's id, forcing
        // VS Code to render them fresh and apply the Collapsed initial state.
        // (Encoding only the node's own active flag is insufficient: a sibling
        // the user manually expanded earlier keeps its id and its expanded UI
        // state across active-module changes.)
        if (node.kind === 'module') {
            const activeKey = this._activeModuleKey ?? '-';
            item.id = `m::${node.filePath}::${node.moduleName}::${activeKey}`;
        } else if (node.kind === 'sub') {
            item.id = `s::${node.filePath}::${node.moduleName}::${node.label}::${node.line ?? 0}`;
        } else if (node.kind === 'xlsm') {
            item.id = `w::${node.isRemote ? node.remoteId ?? '' : node.filePath}`;
        }

        switch (node.kind) {
            case 'xlsm':
                item.iconPath = new vscode.ThemeIcon(node.isRemote ? 'remote' : 'file-code');
                if (node.isRemote) {
                    item.tooltip = `(Live Share) ${node.label}`;
                    item.description = node.remoteRelativeFolder ? `${node.remoteRelativeFolder} (Live Share)` : '(Live Share)';
                    item.contextValue = 'xlsm-remote';
                } else {
                    item.tooltip = node.filePath;
                    item.description = path.relative(
                        vscode.workspace.workspaceFolders?.[0]?.uri.fsPath ?? '',
                        path.dirname(node.filePath),
                    ) || '';
                    item.contextValue = 'xlsm';
                }
                break;

            case 'module':
                item.iconPath = new vscode.ThemeIcon(
                    MODULE_ICONS[node.moduleType ?? 'standard'] ?? 'symbol-module',
                );
                item.description = node.moduleType;
                // Remote modules get a distinct contextValue so host-only menu items don't appear.
                if (node.isRemote) {
                    item.contextValue = `module-remote-${node.moduleType ?? 'standard'}`;
                } else {
                    item.contextValue = `module-${node.moduleType ?? 'standard'}`;
                }
                item.command = {
                    command: 'xlide.openModule',
                    title: 'Open Module',
                    arguments: [node],
                };
                break;

            case 'sub':
                item.iconPath = new vscode.ThemeIcon('symbol-method');
                item.contextValue = node.isRemote ? 'sub-remote' : 'sub';
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
            // Live Share guest sees the host's workbooks; local files are not visible.
            // Use isInGuestSession (session role) rather than isGuest (proxy ready)
            // so we don't briefly render the host's vsls:// URIs as broken local nodes
            // before the shared service proxy has connected.
            if (this._liveShare?.isInGuestSession) {
                if (!this._liveShare.isGuest) {
                    // Proxy not ready yet — render nothing; onDidChange will refresh.
                    return [];
                }
                return this._getRemoteWorkbooks();
            }
            return this._getXlsmFiles();
        }
        if (node.kind === 'xlsm') {
            if (node.isRemote && node.remoteId) {
                return this._getRemoteModules(node.remoteId);
            }
            return this._getModules(node.filePath);
        }
        if (node.kind === 'module') {
            if (node.isRemote && node.remoteId) {
                return this._getRemoteSubs(node.remoteId, node.moduleName!);
            }
            return this._getSubs(node.filePath, node.moduleName!);
        }
        return [];
    }

    private async _getRemoteWorkbooks(): Promise<XlideNode[]> {
        if (!this._liveShare) { return []; }
        try {
            const list = await this._liveShare.guestListWorkbooks();
            return list.map((w) => ({
                kind: 'xlsm' as const,
                label: w.name,
                filePath: '',
                isRemote: true,
                remoteId: w.id,
                remoteRelativeFolder: w.relativeFolder,
            }));
        } catch (err) {
            vscode.window.showErrorMessage(`XLIDE: Failed to list remote workbooks: ${err}`);
            return [];
        }
    }

    private async _getRemoteModules(workbookId: string): Promise<XlideNode[]> {
        if (!this._liveShare) { return []; }
        try {
            const modules = await this._liveShare.guestListModules(workbookId);
            return modules
                .sort((a, b) => {
                    const typeOrder: Record<string, number> = {
                        document: 0, userform: 1, standard: 2, class: 3,
                    };
                    const aOrder = typeOrder[a.type] ?? 4;
                    const bOrder = typeOrder[b.type] ?? 4;
                    if (aOrder !== bOrder) return aOrder - bOrder;
                    return a.name.localeCompare(b.name);
                })
                .map((m) => ({
                    kind: 'module' as const,
                    label: m.name,
                    filePath: '',
                    moduleName: m.name,
                    moduleType: m.type,
                    isRemote: true,
                    remoteId: workbookId,
                }));
        } catch (err) {
            vscode.window.showErrorMessage(`XLIDE: Failed to list remote modules: ${err}`);
            return [];
        }
    }

    private async _getRemoteSubs(workbookId: string, moduleName: string): Promise<XlideNode[]> {
        if (!this._liveShare) { return []; }
        try {
            const subs = await this._liveShare.guestListSubs(workbookId, moduleName);
            return subs.map((s) => ({
                kind: 'sub' as const,
                label: `${s.kind} ${s.name}`,
                filePath: '',
                moduleName,
                line: s.line,
                isRemote: true,
                remoteId: workbookId,
            }));
        } catch {
            return [];
        }
    }

    private async _getXlsmFiles(): Promise<XlideNode[]> {
        const uris = await vscode.workspace.findFiles(
            '**/*.{xlsm,xlsb,xlam}',
            '{**/node_modules/**,**/.venv/**,**/venv/**}',
        );
        return uris
            .filter(uri => uri.scheme === 'file' && !path.basename(uri.fsPath).startsWith('~$'))
            .sort((a, b) => a.fsPath.localeCompare(b.fsPath))
            .map((uri) => {
                let node = this._xlsmNodes.get(uri.fsPath);
                if (!node) {
                    node = { kind: 'xlsm', label: path.basename(uri.fsPath), filePath: uri.fsPath };
                    this._xlsmNodes.set(uri.fsPath, node);
                }
                return node;
            });
    }

    private async _getModules(filePath: string): Promise<XlideNode[]> {
        try {
            let modules = this._modulesListCache.get(filePath);
            if (!modules) {
                modules = await this._bridge.call<Array<{ name: string; type: string }>>(
                    'listModules',
                    { path: filePath },
                );
                this._modulesListCache.set(filePath, modules);
            }
            // Sort: document, userform, standard, class — alphabetical within each group.
            return modules
                .sort((a, b) => {
                    const typeOrder: Record<string, number> = {
                        document: 0, userform: 1, standard: 2, class: 3,
                    };
                    const aOrder = typeOrder[a.type] ?? 4;
                    const bOrder = typeOrder[b.type] ?? 4;
                    if (aOrder !== bOrder) return aOrder - bOrder;
                    return a.name.localeCompare(b.name);
                })
                .map((m) => {
                    const key = `${filePath}::${m.name}`;
                    let node = this._moduleNodes.get(key);
                    if (!node) {
                        node = {
                            kind: 'module',
                            label: m.name,
                            filePath,
                            moduleName: m.name,
                            moduleType: m.type,
                        };
                        this._moduleNodes.set(key, node);
                    }
                    return node;
                });
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
