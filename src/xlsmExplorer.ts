import * as vscode from 'vscode';
import * as path from 'path';
import { PythonBridge } from './pythonBridge';
import { workbookIdentityKey } from './xlideFileSystem';
import type { LiveShareIntegration } from './liveShare';
import { compareVbaModulesForTreeOrder, moduleThemeIconName } from './moduleDisplay';
import { startPerformanceTrace } from './performanceTrace';

export type XlideNodeKind = 'xlsm' | 'module' | 'sub';

const PROTECTION_PROBE_IDLE_DELAY_MS = 2000;

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
    /** Workbook only: VBA project carries a password lock. */
    isPasswordProtected?: boolean;
    /** Workbook only: VBA project carries a digital signature. */
    isSigned?: boolean;
}

export class XlsmExplorer implements vscode.TreeDataProvider<XlideNode>, vscode.Disposable {
    private _emitter = new vscode.EventEmitter<XlideNode | undefined | null | void>();
    readonly onDidChangeTreeData = this._emitter.event;

    private _liveShare: LiveShareIntegration | undefined;
    private _liveShareSubscription: vscode.Disposable | undefined;

    // Stable node references required by treeView.reveal()
    private _xlsmNodes = new Map<string, XlideNode>(); // key: filePath
    private _moduleNodes = new Map<string, XlideNode>(); // key: filePath + '::' + moduleName
    private _xlsmRenderVersions = new Map<string, number>();
    private _moduleRenderVersions = new Map<string, number>();
    private _xlsmFilesCache: XlideNode[] | undefined;
    private _xlsmFilesLoad: Promise<XlideNode[]> | undefined;
    // listModules cache: avoids repeated bridge round-trips while the tree is
    // expanded.  Cleared on refresh() so edits always re-fetch.
    private _modulesListCache = new Map<string, Array<{ name: string; type: string }>>();
    private _modulesListLoads = new Map<string, Promise<Array<{ name: string; type: string }>>>();
    private _subsListCache = new Map<string, Array<{ name: string; kind: string; line: number }>>();
    private _subsListLoads = new Map<string, Promise<Array<{ name: string; kind: string; line: number }>>>();
    // Protection-state cache: {isPasswordProtected, isSigned} per workbook path.
    // Loaded lazily after tree expansion has gone idle; cleared on refresh().
    private _protectionCache = new Map<string, { isPasswordProtected: boolean; isSigned: boolean }>();
    private _protectionLoads = new Map<string, Promise<void>>();
    private _protectionTimers = new Map<string, ReturnType<typeof setTimeout>>();
    // Accordion: only one module node is expanded at a time.
    private _activeModuleKey: string | undefined;
    private _activeWorkbookKey: string | undefined;
    private _setupComplete = false;

    constructor(
        private readonly _bridge: PythonBridge,
        private readonly _out?: vscode.OutputChannel,
    ) {}

    setLiveShare(liveShare: LiveShareIntegration): void {
        this._liveShare = liveShare;
        this._liveShareSubscription?.dispose();
        this._liveShareSubscription = liveShare.onDidChange(() => this.refresh());
    }

    dispose(): void {
        this._clearProtectionTimers();
        this._liveShareSubscription?.dispose();
        this._liveShareSubscription = undefined;
        this._emitter.dispose();
    }

    setSetupComplete(setupComplete: boolean): void {
        if (this._setupComplete === setupComplete) {
            return;
        }
        this._setupComplete = setupComplete;
        this.refresh();
    }

    refresh(): void {
        this._xlsmNodes.clear();
        this._moduleNodes.clear();
        this._xlsmRenderVersions.clear();
        this._moduleRenderVersions.clear();
        this._xlsmFilesCache = undefined;
        this._xlsmFilesLoad = undefined;
        this._modulesListCache.clear();
        this._modulesListLoads.clear();
        this._subsListCache.clear();
        this._subsListLoads.clear();
        this._protectionCache.clear();
        this._protectionLoads.clear();
        this._clearProtectionTimers();
        this._emitter.fire();
    }

    /**
     * Refresh just the children of a single module node (i.e. its sub list)
     * without collapsing the tree or clearing other caches. If the module node
     * isn't loaded yet, this is a no-op.
     */
    refreshModuleSubs(filePath: string, moduleName: string): void {
        const key = moduleNodeKey(filePath, moduleName);
        this._subsListCache.delete(key);
        this._subsListLoads.delete(key);
        const node = this._moduleNodes.get(key);
        if (node) {
            this._emitter.fire(node);
        }
    }

    /** Required by treeView.reveal() — walks xlsm -> module -> sub. */
    getParent(node: XlideNode): XlideNode | undefined {
        if (node.kind === 'module') {
            return this._xlsmNodes.get(node.filePath);
        }
        if (node.kind === 'sub') {
            return this._moduleNodes.get(moduleNodeKey(node.filePath, node.moduleName ?? ''));
        }
        return undefined;
    }

    /** Returns the cached module node, if the tree has loaded it. */
    getModuleNode(filePath: string, moduleName: string): XlideNode | undefined {
        return this._moduleNodes.get(moduleNodeKey(filePath, moduleName));
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
        const key = moduleNodeKey(filePath, moduleName);
        if (this._activeModuleKey === key) { return; }
        const previousKey = this._activeModuleKey;
        const previousWorkbookKey = this._activeWorkbookKey;
        const nextWorkbookKey = workbookNodeKey(filePath);
        this._activeModuleKey = key;
        this._activeWorkbookKey = nextWorkbookKey;
        this._refreshModuleExpansion(previousKey);
        this._refreshModuleExpansion(key);
        if (previousWorkbookKey !== nextWorkbookKey) {
            this._refreshNonActiveWorkbookExpansions(nextWorkbookKey);
        }
    }

    /**
     * Clear the active module (e.g. when the last XLIDE editor closes) so that
     * every module under every workbook collapses.
     */
    clearActiveModule(): void {
        if (this._activeModuleKey === undefined) { return; }
        const previousKey = this._activeModuleKey;
        this._activeModuleKey = undefined;
        this._activeWorkbookKey = undefined;
        this._refreshModuleExpansion(previousKey);
        this._refreshNonActiveWorkbookExpansions(undefined);
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
            moduleNodeKey(node.filePath, node.moduleName ?? '') === this._activeModuleKey;
        const isActiveWorkbook =
            node.kind === 'xlsm' &&
            !node.isRemote &&
            workbookNodeKey(node.filePath) === this._activeWorkbookKey;
        const item = new vscode.TreeItem(
            node.label,
            node.kind === 'sub'
                ? vscode.TreeItemCollapsibleState.None
                : isActiveModule
                    ? vscode.TreeItemCollapsibleState.Expanded
                    : isActiveWorkbook
                        ? vscode.TreeItemCollapsibleState.Expanded
                    : vscode.TreeItemCollapsibleState.Collapsed,
        );

        if (node.kind === 'module') {
            const key = moduleNodeKey(node.filePath, node.moduleName ?? '');
            const version = this._moduleRenderVersions.get(key) ?? 0;
            item.id = `m::${key}::${version}`;
        } else if (node.kind === 'sub') {
            item.id = `s::${node.filePath}::${node.moduleName}::${node.label}::${node.line ?? 0}`;
        } else if (node.kind === 'xlsm') {
            const key = node.isRemote ? node.remoteId ?? '' : workbookNodeKey(node.filePath);
            const version = this._xlsmRenderVersions.get(key) ?? 0;
            item.id = `w::${key}::${version}`;
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
                    // Append protection/signature badges once known.
                    const badges: string[] = [];
                    if (node.isPasswordProtected) { badges.push('locked'); }
                    if (node.isSigned) { badges.push('signed'); }
                    if (badges.length > 0) {
                        const tag = `[${badges.join(', ')}]`;
                        item.description = item.description ? `${item.description}  ${tag}` : tag;
                        const tip = new vscode.MarkdownString(node.filePath);
                        if (node.isPasswordProtected) {
                            tip.appendMarkdown('\n\n$(lock) VBA project is password-protected');
                        }
                        if (node.isSigned) {
                            tip.appendMarkdown('\n\n$(shield) VBA project is digitally signed (edits will invalidate the signature)');
                        }
                        tip.supportThemeIcons = true;
                        item.tooltip = tip;
                    }
                }
                break;

            case 'module':
                item.iconPath = new vscode.ThemeIcon(moduleThemeIconName(node.moduleType));
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
        const trace = startPerformanceTrace('tree.getChildren', node?.kind ?? 'root');
        try {
            const result = await this._getChildren(node);
            trace.end('ok', node?.kind ?? 'root');
            return result;
        } catch (err) {
            trace.end('failed', node?.kind ?? 'root');
            throw err;
        }
    }

    private async _getChildren(node?: XlideNode): Promise<XlideNode[]> {
        if (!node) {
            if (!this._setupComplete) {
                return [];
            }
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
                .sort(compareVbaModulesForTreeOrder)
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
        } catch (err) {
            vscode.window.showErrorMessage(`XLIDE: Failed to list procedures in remote module "${moduleName}": ${err}`);
            return [];
        }
    }

    private async _getXlsmFiles(): Promise<XlideNode[]> {
        if (this._xlsmFilesCache) {
            return this._xlsmFilesCache;
        }
        if (this._xlsmFilesLoad) {
            return this._xlsmFilesLoad;
        }
        const load = this._loadXlsmFiles();
        this._xlsmFilesLoad = load;
        try {
            const nodes = await load;
            this._xlsmFilesCache = nodes;
            return nodes;
        } finally {
            if (this._xlsmFilesLoad === load) {
                this._xlsmFilesLoad = undefined;
            }
        }
    }

    private async _loadXlsmFiles(): Promise<XlideNode[]> {
        const uris = await vscode.workspace.findFiles(
            '**/*.{xlsm,xlsb,xlam}',
            '{**/node_modules/**,**/.venv/**,**/venv/**}',
        );
        return uris
            .filter(uri => uri.scheme === 'file' && !fileNameForDisplay(uri.fsPath).startsWith('~$'))
            .sort((a, b) => a.fsPath.localeCompare(b.fsPath))
            .map((uri) => {
                let node = this._xlsmNodes.get(uri.fsPath);
                if (!node) {
                    node = { kind: 'xlsm', label: fileNameForDisplay(uri.fsPath), filePath: uri.fsPath };
                    this._xlsmNodes.set(uri.fsPath, node);
                }
                return node;
            });
    }

    private async _getModules(filePath: string): Promise<XlideNode[]> {
        this._cancelProtectionTimer(filePath);
        try {
            const cacheKey = workbookNodeKey(filePath);
            let modules = this._modulesListCache.get(cacheKey);
            if (!modules) {
                let load = this._modulesListLoads.get(cacheKey);
                if (!load) {
                    load = this._bridge.call<Array<{ name: string; type: string }>>(
                        'listModules',
                        { path: filePath },
                    );
                    this._modulesListLoads.set(cacheKey, load);
                    load.then(
                        () => {
                            if (this._modulesListLoads.get(cacheKey) === load) {
                                this._modulesListLoads.delete(cacheKey);
                            }
                        },
                        () => {
                            if (this._modulesListLoads.get(cacheKey) === load) {
                                this._modulesListLoads.delete(cacheKey);
                            }
                        },
                    );
                }
                // Sort once, on a copy, before caching: the cached list is then
                // already in tree order, so re-renders (cache hits) skip the sort
                // and never mutate the shared cache.
                modules = [...await load].sort(compareVbaModulesForTreeOrder);
                this._modulesListCache.set(cacheKey, modules);
            }
            const nodes = modules
                .map((m) => {
                    const key = moduleNodeKey(filePath, m.name);
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
            this._scheduleProtectionLoad(filePath);
            return nodes;
        } catch (err) {
            vscode.window.showErrorMessage(`XLIDE: Failed to list modules in "${fileNameForDisplay(filePath)}": ${err}`);
            return [];
        }
    }

    /**
     * Lazily fetch the workbook's VBA protection/signature state and, once
     * known, stamp it onto the cached xlsm node and re-render so the tree item
     * shows the locked/signed badge. Best-effort: failures are ignored.
     */
    private async _loadProtection(filePath: string): Promise<void> {
        if (this._protectionCache.has(filePath)) { return; }
        const existing = this._protectionLoads.get(filePath);
        if (existing) {
            await existing;
            return;
        }
        const load = (async () => {
            try {
                const info = await this._bridge.call<{ isPasswordProtected: boolean; isSigned: boolean }>(
                    'getProtectionInfo',
                    { path: filePath },
                );
                this._protectionCache.set(filePath, info);
                const node = this._xlsmNodes.get(filePath);
                if (node) {
                    node.isPasswordProtected = info.isPasswordProtected;
                    node.isSigned = info.isSigned;
                    this._emitter.fire(node);
                }
            } catch (err) {
                // Badge is best-effort; log the probe failure without surfacing it.
                this._out?.appendLine(`[xlsmExplorer] Protection probe failed for "${fileNameForDisplay(filePath)}": ${err}`);
            } finally {
                this._protectionLoads.delete(filePath);
            }
        })();
        this._protectionLoads.set(filePath, load);
        await load;
    }

    private async _getSubs(filePath: string, moduleName: string): Promise<XlideNode[]> {
        this._cancelProtectionTimer(filePath);
        const cacheKey = moduleNodeKey(filePath, moduleName);
        try {
            let subs = this._subsListCache.get(cacheKey);
            if (!subs) {
                let load = this._subsListLoads.get(cacheKey);
                if (!load) {
                    load = this._bridge.call<Array<{ name: string; kind: string; line: number }>>(
                        'listSubs',
                        { path: filePath, module: moduleName },
                    );
                    this._subsListLoads.set(cacheKey, load);
                    load.then(
                        () => {
                            if (this._subsListLoads.get(cacheKey) === load) {
                                this._subsListLoads.delete(cacheKey);
                            }
                        },
                        () => {
                            if (this._subsListLoads.get(cacheKey) === load) {
                                this._subsListLoads.delete(cacheKey);
                            }
                        },
                    );
                }
                subs = await load;
                this._subsListCache.set(cacheKey, subs);
            }
            const nodes = subs.map((s) => ({
                kind: 'sub' as const,
                label: `${s.kind} ${s.name}`,
                filePath,
                moduleName,
                line: s.line,
            }));
            this._scheduleProtectionLoad(filePath);
            return nodes;
        } catch (err) {
            vscode.window.showErrorMessage(`XLIDE: Failed to list procedures in "${moduleName}" (${fileNameForDisplay(filePath)}): ${err}`);
            return [];
        }
    }

    private _refreshModuleExpansion(key: string | undefined): void {
        if (!key) { return; }
        this._moduleRenderVersions.set(key, (this._moduleRenderVersions.get(key) ?? 0) + 1);
        const node = this._moduleNodes.get(key);
        if (node) {
            this._emitter.fire(node);
        }
    }

    private _scheduleProtectionLoad(filePath: string): void {
        if (this._protectionCache.has(filePath) || this._protectionLoads.has(filePath)) {
            return;
        }
        if (this._protectionTimers.has(filePath)) {
            return;
        }
        const timer = setTimeout(() => {
            this._protectionTimers.delete(filePath);
            void this._loadProtection(filePath);
        }, PROTECTION_PROBE_IDLE_DELAY_MS);
        (timer as unknown as { unref?: () => void }).unref?.();
        this._protectionTimers.set(filePath, timer);
    }

    private _cancelProtectionTimer(filePath: string): void {
        const timer = this._protectionTimers.get(filePath);
        if (!timer) {
            return;
        }
        clearTimeout(timer);
        this._protectionTimers.delete(filePath);
    }

    private _clearProtectionTimers(): void {
        for (const timer of this._protectionTimers.values()) {
            clearTimeout(timer);
        }
        this._protectionTimers.clear();
    }

    private _refreshNonActiveWorkbookExpansions(activeWorkbookKey: string | undefined): void {
        for (const node of this._xlsmNodes.values()) {
            if (node.isRemote) {
                continue;
            }
            const key = workbookNodeKey(node.filePath);
            if (key === activeWorkbookKey) {
                continue;
            }
            this._xlsmRenderVersions.set(key, (this._xlsmRenderVersions.get(key) ?? 0) + 1);
            this._emitter.fire(node);
        }
    }
}

function fileNameForDisplay(filePath: string): string {
    return filePath.split(/[\\/]/).filter(Boolean).pop() ?? filePath;
}

function moduleNodeKey(filePath: string, moduleName: string): string {
    return `${filePath}::${moduleName}`;
}

function workbookNodeKey(filePath: string): string {
    return workbookIdentityKey(filePath);
}
