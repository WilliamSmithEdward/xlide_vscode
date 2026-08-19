import * as vscode from 'vscode';
import * as path from 'path';
import { WorkbookEngine } from './workbookEngine';
import { workbookIdentityKey } from './xlideFileSystem';
import { compareVbaModulesForTreeOrder, moduleThemeIconName } from './moduleDisplay';
import { containerAppNameForPath, containerContextValue, isReadOnlyContainerPath, MACRO_CONTAINER_GLOB } from './macroContainerUi';
import { hasPendingAgentReview } from './xlideAgentDiff';
import { startPerformanceTrace } from './performanceTrace';

export type XlideNodeKind = 'xlsm' | 'module' | 'sub' | 'loadError';

const PROTECTION_PROBE_IDLE_DELAY_MS = 2000;

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
    /** Workbook only: VBA project carries a password lock. */
    isPasswordProtected?: boolean;
    /** loadError only: the failure message shown in the tooltip. */
    errorMessage?: string;
    /** Workbook only: VBA project carries a digital signature. */
    isSigned?: boolean;
}

export class XlsmExplorer implements vscode.TreeDataProvider<XlideNode>, vscode.Disposable {
    private _emitter = new vscode.EventEmitter<XlideNode | undefined | null | void>();
    readonly onDidChangeTreeData = this._emitter.event;


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
    // Bumped on every refresh(). An in-flight load captured before a refresh must
    // not write its now-stale result into the freshly-cleared cache (which would
    // leave a just-added module invisible until the next refresh).
    private _generation = 0;
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

    constructor(
        private readonly _bridge: WorkbookEngine,
        private readonly _out?: vscode.OutputChannel,
    ) {}

    dispose(): void {
        this._clearProtectionTimers();
        this._emitter.dispose();
    }

    refresh(): void {
        this._generation++;
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

    /** Required by treeView.reveal() - walks xlsm -> module -> sub. */
    getParent(node: XlideNode): XlideNode | undefined {
        if (node.kind === 'module') {
            return this._xlsmNodes.get(node.filePath);
        }
        if (node.kind === 'sub') {
            return this._moduleNodes.get(moduleNodeKey(node.filePath, node.moduleName ?? ''));
        }
        if (node.kind === 'loadError') {
            return node.moduleName
                ? this._moduleNodes.get(moduleNodeKey(node.filePath, node.moduleName))
                : this._xlsmNodes.get(node.filePath);
        }
        return undefined;
    }

    /**
     * Retry a failed workbook/module listing from its in-tree "click to retry"
     * placeholder. A transient failure (e.g. Excel briefly holding an exclusive
     * lock on the file during open/save) must never permanently brick the node:
     * VS Code caches resolved children until the node is re-fired, so without
     * this the tree would stay broken until a full refresh or window reload.
     */
    retryLoad(node: XlideNode): void {
        if (node.kind !== 'loadError') {
            return;
        }
        if (node.moduleName) {
            this.refreshModuleSubs(node.filePath, node.moduleName);
            return;
        }
        const key = workbookNodeKey(node.filePath);
        this._modulesListCache.delete(key);
        this._modulesListLoads.delete(key);
        const workbook = this._xlsmNodes.get(node.filePath);
        if (workbook) {
            this._emitter.fire(workbook);
        } else {
            this._emitter.fire();
        }
    }

    private _loadErrorNode(filePath: string, moduleName: string | undefined, err: unknown): XlideNode {
        return {
            kind: 'loadError',
            label: 'Load failed - click to retry',
            filePath,
            moduleName,
            errorMessage: String(err),
        };
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
            // Only when actually switching away from a previously-active workbook
            // (not the first activation, which has nothing to collapse): bump the
            // now-active workbook's render id so it re-renders Expanded, then
            // refresh from the root so VS Code rebuilds the workbook list and
            // actually applies the new collapsed/expanded states. Firing the
            // individual workbook nodes above only refreshes them in place, which
            // does not reliably collapse an already-expanded workbook.
            if (previousWorkbookKey !== undefined) {
                this._xlsmRenderVersions.set(
                    nextWorkbookKey,
                    (this._xlsmRenderVersions.get(nextWorkbookKey) ?? 0) + 1,
                );
                this._emitter.fire();
            }
        }
    }

    /**
     * Clears the forced-expand state for a workbook the user manually collapsed,
     * so a later refresh does not re-stamp it Expanded and spring it back open.
     */
    notifyWorkbookCollapsed(filePath: string): void {
        if (this._activeWorkbookKey === workbookNodeKey(filePath)) {
            this._activeWorkbookKey = undefined;
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
            moduleNodeKey(node.filePath, node.moduleName ?? '') === this._activeModuleKey;
        const isActiveWorkbook =
            node.kind === 'xlsm' &&
            workbookNodeKey(node.filePath) === this._activeWorkbookKey;
        const item = new vscode.TreeItem(
            node.label,
            node.kind === 'sub' || node.kind === 'loadError'
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
            const key = workbookNodeKey(node.filePath);
            const version = this._xlsmRenderVersions.get(key) ?? 0;
            item.id = `w::${key}::${version}`;
        }

        switch (node.kind) {
            case 'xlsm':
                item.iconPath = new vscode.ThemeIcon('file-code');
                item.tooltip = node.filePath;
                item.description = path.relative(
                    vscode.workspace.workspaceFolders?.[0]?.uri.fsPath ?? '',
                    path.dirname(node.filePath),
                ) || '';
                item.contextValue = containerContextValue(node.filePath);
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
                break;

            case 'module':
                item.iconPath = new vscode.ThemeIcon(moduleThemeIconName(node.moduleType));
                item.description = node.moduleType;
                // '-ro' keeps rename/delete menus off read-only containers.
                item.contextValue = `module-${node.moduleType ?? 'standard'}${
                    isReadOnlyContainerPath(node.filePath) ? '-ro' : ''}`;
                if (hasPendingAgentReview(node.filePath, node.moduleName ?? '')) {
                    // An agent wrote this module and nobody has kept or
                    // reverted it yet; the badge keeps the review reachable.
                    item.description = `${node.moduleType} ● agent edit`;
                    item.contextValue += '-agent-pending';
                    item.iconPath = new vscode.ThemeIcon(
                        moduleThemeIconName(node.moduleType),
                        new vscode.ThemeColor('gitDecoration.modifiedResourceForeground'),
                    );
                }
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

            case 'loadError':
                item.iconPath = new vscode.ThemeIcon('warning', new vscode.ThemeColor('problemsWarningIcon.foreground'));
                item.contextValue = 'loadError';
                item.tooltip = `${node.errorMessage ?? 'The file could not be read.'}\n\nIf the file is open in ${containerAppNameForPath(node.filePath)}, close it (or wait for the save to finish) and click to retry.`;
                item.command = {
                    command: 'xlide.retryExplorerLoad',
                    title: 'Retry',
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
            MACRO_CONTAINER_GLOB,
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
            let loadGeneration: number | undefined;
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
                const generation = this._generation;
                loadGeneration = generation;
                modules = [...await load].sort(compareVbaModulesForTreeOrder);
                // Only cache when no refresh() raced this load to completion;
                // otherwise the post-refresh render will re-fetch the fresh list.
                if (this._generation === generation) {
                    this._modulesListCache.set(cacheKey, modules);
                }
            }
            // Only populate the shared node-identity map when this render is not a
            // stale fresh-load that a refresh() raced to completion; otherwise it
            // would re-insert old-generation node identities into the cleared map.
            const populateNodeMap = loadGeneration === undefined || this._generation === loadGeneration;
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
                        if (populateNodeMap) {
                            this._moduleNodes.set(key, node);
                        }
                    }
                    return node;
                });
            this._scheduleProtectionLoad(filePath);
            return nodes;
        } catch (err) {
            vscode.window.showErrorMessage(`XLIDE: Failed to list modules in "${fileNameForDisplay(filePath)}": ${err}`);
            // Return a retry placeholder, never [] - VS Code caches resolved
            // children, so an empty result would leave the workbook permanently
            // empty after a transient failure (e.g. Excel holding the file).
            return [this._loadErrorNode(filePath, undefined, err)];
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
                // Only cache when no refresh() raced this load to completion;
                // otherwise the stale sub list would poison the freshly-cleared
                // cache (mirrors the generation guard in _getModules).
                const generation = this._generation;
                subs = await load;
                if (this._generation === generation) {
                    this._subsListCache.set(cacheKey, subs);
                }
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
            return [this._loadErrorNode(filePath, moduleName, err)];
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
