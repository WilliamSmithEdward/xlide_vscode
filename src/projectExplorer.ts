import * as vscode from 'vscode';
import * as path from 'path';
import { ProjectEngine } from './projectEngine';
import { projectIdentityKey } from './xlideFileSystem';
import { compareVbaModulesForTreeOrder, moduleThemeIconName } from './moduleDisplay';
import { buildFolderTree, folderPathChain, type FolderTree, type FolderTreeFolder } from './folderTree';
import type { XlideExplorerView } from './globalSettings';
import { containerAppNameForPath, containerContextValue, isVb6ProjectPath } from './macroContainerUi';
import { findMacroContainerFiles } from './macroContainerDiscovery';
import { hasPendingAgentReview } from './xlideAgentDiff';
import { startPerformanceTrace } from './performanceTrace';

export type XlideNodeKind = 'project' | 'folder' | 'module' | 'designer' | 'sub' | 'loadError';

export type { XlideExplorerView } from './globalSettings';

const PROTECTION_PROBE_IDLE_DELAY_MS = 2000;

export interface XlideNode {
    kind: XlideNodeKind;
    label: string;
    /** Absolute path to the .xlsm file. */
    filePath: string;
    /** Module name (for 'module' and 'sub' nodes). */
    moduleName?: string;
    /** Module type: 'standard' | 'class' | 'document' | 'userform' | 'accessform' | 'accessreport', or a VB6-only kind. */
    moduleType?: string;
    /**
     * The module's own file, when the container's modules are files (a VB6
     * project): opening the node opens this file, not a virtual document.
     */
    moduleFilePath?: string;
    /** 1-based line number of the procedure (for 'sub' nodes). */
    line?: number;
    /**
     * The module's `@Folder` annotation, normalized; absent puts the module at
     * the project's root. Read on 'module' nodes, and the whole dotted path on
     * a 'folder' node - which is the annotation that makes that folder.
     */
    folder?: string;
    /** folder only: modules in this folder and every folder under it. */
    moduleCount?: number;
    /** Workbook only: VBA project carries a password lock. */
    isPasswordProtected?: boolean;
    /** loadError only: the failure message shown in the tooltip. */
    errorMessage?: string;
    /** Workbook only: VBA project carries a digital signature. */
    isSigned?: boolean;
}

export class ProjectExplorer implements vscode.TreeDataProvider<XlideNode>, vscode.Disposable {
    private _emitter = new vscode.EventEmitter<XlideNode | undefined | null | void>();
    readonly onDidChangeTreeData = this._emitter.event;


    // Stable node references required by treeView.reveal()
    private _projectNodes = new Map<string, XlideNode>(); // key: filePath
    private _moduleNodes = new Map<string, XlideNode>(); // key: filePath + '::' + moduleName
    private _projectRenderVersions = new Map<string, number>();
    private _moduleRenderVersions = new Map<string, number>();
    private _projectFilesCache: XlideNode[] | undefined;
    private _projectFilesLoad: Promise<XlideNode[]> | undefined;
    // listModules cache: avoids repeated bridge round-trips while the tree is
    // expanded.  Cleared on refresh() so edits always re-fetch.
    private _modulesListCache = new Map<string, Array<{ name: string; type: string; filePath?: string; folder?: string }>>();
    private _modulesListLoads = new Map<string, Promise<Array<{ name: string; type: string; filePath?: string; folder?: string }>>>();
    // Bumped on every refresh(). An in-flight load captured before a refresh must
    // not write its now-stale result into the freshly-cleared cache (which would
    // leave a just-added module invisible until the next refresh).
    private _generation = 0;
    private _subsListCache = new Map<string, Array<{ name: string; kind: string; line: number }>>();
    private _subsListLoads = new Map<string, Promise<Array<{ name: string; kind: string; line: number }>>>();
    // The drawn sub/designer rows of a module, kept so reveal() can name one.
    private _subNodes = new Map<string, XlideNode[]>();
    // Protection-state cache: {isPasswordProtected, isSigned} per project path.
    // Loaded lazily after tree expansion has gone idle; cleared on refresh().
    private _protectionCache = new Map<string, { isPasswordProtected: boolean; isSigned: boolean }>();
    private _protectionLoads = new Map<string, Promise<void>>();
    private _protectionTimers = new Map<string, ReturnType<typeof setTimeout>>();
    // Accordion: only one module node is expanded at a time.
    private _activeModuleKey: string | undefined;
    private _activeProjectKey: string | undefined;
    // Folder layout. The tree is derived from the module list, so it is cached
    // beside it and thrown away by the same refresh().
    private _view: XlideExplorerView = 'tree';
    private _folderNodes = new Map<string, XlideNode>(); // key: filePath + '::' + folder
    private _folderRenderVersions = new Map<string, number>();
    private _folderTrees = new Map<string, FolderTree<XlideNode>>();
    // Folders the editor's own module opened, and the ones the user opened or
    // shut by hand, which outrank it until the attention genuinely moves.
    private _openFolderKeys = new Set<string>();
    private _manualFolderStates = new Map<string, boolean>();
    private _activeFolderChain: string[] = [];
    // What an open editor says a module's @Folder annotation is, which outranks
    // the listing until that editor closes. Survives refresh() on purpose.
    private _editorFolders = new Map<string, string | undefined>();

    constructor(
        private readonly _bridge: ProjectEngine,
        private readonly _out?: vscode.OutputChannel,
    ) {}

    dispose(): void {
        this._clearProtectionTimers();
        this._emitter.dispose();
    }

    /** The layout the tree draws: the flat module list, or the folder layout. */
    get view(): XlideExplorerView {
        return this._view;
    }

    /**
     * Switch layouts. The two draw the same modules in different places, so the
     * whole tree is rebuilt; the folders the editor opened are kept, so
     * switching to the folder layout lands on the module being edited.
     */
    setView(view: XlideExplorerView): void {
        if (this._view === view) { return; }
        this._view = view;
        this.refresh();
    }

    refresh(): void {
        this._generation++;
        this._projectNodes.clear();
        this._moduleNodes.clear();
        this._folderNodes.clear();
        this._folderTrees.clear();
        this._projectRenderVersions.clear();
        this._moduleRenderVersions.clear();
        this._folderRenderVersions.clear();
        this._projectFilesCache = undefined;
        this._projectFilesLoad = undefined;
        this._modulesListCache.clear();
        this._modulesListLoads.clear();
        this._subsListCache.clear();
        this._subsListLoads.clear();
        this._subNodes.clear();
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
        this._subNodes.delete(key);
        const node = this._moduleNodes.get(key);
        if (node) {
            this._emitter.fire(node);
        }
    }

    /** Required by treeView.reveal() - walks xlsm -> (folder) -> module -> sub. */
    getParent(node: XlideNode): XlideNode | undefined {
        if (node.kind === 'folder') {
            const parent = parentFolderPath(node.folder ?? '');
            return parent
                ? this._folderNodes.get(folderNodeKey(node.filePath, parent))
                : this._projectNodes.get(node.filePath);
        }
        if (node.kind === 'module') {
            if (this._view === 'folders' && node.folder) {
                return this._folderNodes.get(folderNodeKey(node.filePath, node.folder));
            }
            return this._projectNodes.get(node.filePath);
        }
        if (node.kind === 'sub') {
            return this._moduleNodes.get(moduleNodeKey(node.filePath, node.moduleName ?? ''));
        }
        if (node.kind === 'loadError') {
            return node.moduleName
                ? this._moduleNodes.get(moduleNodeKey(node.filePath, node.moduleName))
                : this._projectNodes.get(node.filePath);
        }
        return undefined;
    }

    /**
     * Retry a failed project/module listing from its in-tree "click to retry"
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
        const key = projectNodeKey(node.filePath);
        this._modulesListCache.delete(key);
        this._modulesListLoads.delete(key);
        const project = this._projectNodes.get(node.filePath);
        if (project) {
            this._emitter.fire(project);
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
    getProjectNode(filePath: string): XlideNode | undefined {
        return this._projectNodes.get(filePath);
    }

    /**
     * The row for one procedure, named the way the tree labels it ("Sub Post",
     * "Property Get Name"). Undefined until the module's procedures have been
     * listed, and while an unsaved rename has the editor and the container
     * calling the same procedure different things.
     */
    getProcedureNode(filePath: string, moduleName: string, label: string): XlideNode | undefined {
        const wanted = label.toLowerCase();
        return this._subNodes
            .get(moduleNodeKey(filePath, moduleName))
            ?.find((node) => node.kind === 'sub' && node.label.toLowerCase() === wanted);
    }

    /**
     * Accordion-expand the given module and collapse all sibling module nodes
     * under the same project. Safe to call before the tree has loaded.
     */
    setActiveModule(filePath: string, moduleName: string): void {
        const key = moduleNodeKey(filePath, moduleName);
        if (this._activeModuleKey === key) { return; }
        const previousKey = this._activeModuleKey;
        const previousProjectKey = this._activeProjectKey;
        const nextProjectKey = projectNodeKey(filePath);
        this._activeModuleKey = key;
        this._activeProjectKey = nextProjectKey;
        this._refreshModuleExpansion(previousKey);
        this._refreshModuleExpansion(key);
        let refreshFromRoot = this._followFoldersTo(filePath, moduleName);
        if (previousProjectKey !== nextProjectKey) {
            this._refreshNonActiveProjectExpansions(nextProjectKey);
            // Only when actually switching away from a previously-active project
            // (not the first activation, which has nothing to collapse): bump the
            // now-active project's render id so it re-renders Expanded, then
            // refresh from the root so VS Code rebuilds the project list and
            // actually applies the new collapsed/expanded states. Firing the
            // individual project nodes above only refreshes them in place, which
            // does not reliably collapse an already-expanded project.
            if (previousProjectKey !== undefined) {
                this._projectRenderVersions.set(
                    nextProjectKey,
                    (this._projectRenderVersions.get(nextProjectKey) ?? 0) + 1,
                );
                refreshFromRoot = true;
            }
        }
        if (refreshFromRoot) {
            this._emitter.fire();
        }
    }

    /**
     * Open the folders on the way to the module being edited and fold the
     * project's others. A folder the user opened or shut by hand outranks this
     * until the attention genuinely moves, which is what a changed chain means:
     * moving between two modules of one folder leaves the tree alone.
     *
     * Returns whether the tree has to be rebuilt from the root, which is what
     * makes an already-expanded folder actually collapse.
     */
    private _followFoldersTo(filePath: string, moduleName: string): boolean {
        const module = this._moduleNodes.get(moduleNodeKey(filePath, moduleName));
        if (!module) {
            // The tree has not loaded this module, so its folder is unknown and
            // guessing "no folder" would fold the tree the editor is inside.
            return false;
        }
        const chain = folderPathChain(module.folder).map((step) => folderNodeKey(filePath, step));
        if (sameOrder(chain, this._activeFolderChain)) {
            return false;
        }
        const touched = new Set([
            ...this._openFolderKeys,
            ...this._manualFolderStates.keys(),
            ...chain,
        ]);
        this._activeFolderChain = chain;
        this._openFolderKeys = new Set(chain);
        this._manualFolderStates.clear();
        return this._bumpFolderVersions(touched);
    }

    /** Whether a folder node renders expanded, by hand or by the editor. */
    private _folderIsOpen(node: XlideNode): boolean {
        const key = folderNodeKey(node.filePath, node.folder ?? '');
        return this._manualFolderStates.get(key) ?? this._openFolderKeys.has(key);
    }

    /** Returns the cached folder node, if the tree has drawn it. */
    getFolderNode(filePath: string, folder: string): XlideNode | undefined {
        return this._folderNodes.get(folderNodeKey(filePath, folder));
    }

    /**
     * Record a folder the user opened or shut themselves. It holds that state
     * until the editor moves to a module in a different folder.
     */
    notifyFolderExpansion(node: XlideNode, expanded: boolean): void {
        if (node.kind !== 'folder') { return; }
        this._manualFolderStates.set(folderNodeKey(node.filePath, node.folder ?? ''), expanded);
    }

    /**
     * Fold every folder, in every project. The last editor closing leaves no
     * module for the tree to follow, so it goes back to its resting shape.
     */
    collapseAllFolders(): void {
        if (this._openFolderKeys.size === 0 && this._manualFolderStates.size === 0) { return; }
        const touched = new Set([...this._openFolderKeys, ...this._manualFolderStates.keys()]);
        this._activeFolderChain = [];
        this._openFolderKeys.clear();
        this._manualFolderStates.clear();
        if (this._bumpFolderVersions(touched)) {
            this._emitter.fire();
        }
    }

    /** New render ids for the folders whose expansion changed. */
    private _bumpFolderVersions(keys: Iterable<string>): boolean {
        let any = false;
        for (const key of keys) {
            this._folderRenderVersions.set(key, (this._folderRenderVersions.get(key) ?? 0) + 1);
            any = true;
        }
        return any;
    }

    /**
     * Clears the forced-expand state for a project the user manually collapsed,
     * so a later refresh does not re-stamp it Expanded and spring it back open.
     */
    notifyProjectCollapsed(filePath: string): void {
        if (this._activeProjectKey === projectNodeKey(filePath)) {
            this._activeProjectKey = undefined;
        }
    }

    /**
     * Eagerly loads and caches the root xlsm nodes without waiting for the tree
     * to expand them. Returns the first node (if any) so callers can auto-reveal.
     */
    async warmProjectCache(): Promise<XlideNode | undefined> {
        const nodes = await this._getProjectFiles();
        return nodes[0];
    }

    getTreeItem(node: XlideNode): vscode.TreeItem {
        const isActiveModule =
            node.kind === 'module' &&
            moduleNodeKey(node.filePath, node.moduleName ?? '') === this._activeModuleKey;
        const isActiveProject =
            node.kind === 'project' &&
            projectNodeKey(node.filePath) === this._activeProjectKey;
        const isOpenFolder = node.kind === 'folder' && this._folderIsOpen(node);
        const item = new vscode.TreeItem(
            node.label,
            node.kind === 'sub' || node.kind === 'designer' || node.kind === 'loadError'
                ? vscode.TreeItemCollapsibleState.None
                : isActiveModule || isActiveProject || isOpenFolder
                    ? vscode.TreeItemCollapsibleState.Expanded
                    : vscode.TreeItemCollapsibleState.Collapsed,
        );

        if (node.kind === 'module') {
            const key = moduleNodeKey(node.filePath, node.moduleName ?? '');
            const version = this._moduleRenderVersions.get(key) ?? 0;
            item.id = `m::${key}::${version}`;
        } else if (node.kind === 'sub') {
            item.id = `s::${node.filePath}::${node.moduleName}::${node.label}::${node.line ?? 0}`;
        } else if (node.kind === 'designer') {
            item.id = `d::${node.filePath}::${node.moduleName}`;
        } else if (node.kind === 'project') {
            const key = projectNodeKey(node.filePath);
            const version = this._projectRenderVersions.get(key) ?? 0;
            item.id = `w::${key}::${version}`;
        } else if (node.kind === 'folder') {
            const key = folderNodeKey(node.filePath, node.folder ?? '');
            const version = this._folderRenderVersions.get(key) ?? 0;
            item.id = `f::${key}::${version}`;
        }

        switch (node.kind) {
            case 'project':
                // A VB6 project is a manifest over files, and its icon says so.
                item.iconPath = new vscode.ThemeIcon(isVb6ProjectPath(node.filePath) ? 'project' : 'file-code');
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

            case 'folder':
                item.iconPath = new vscode.ThemeIcon('folder');
                item.contextValue = 'folder';
                // The annotation that makes the folder, so a nested one reads
                // as the whole path rather than just its last segment.
                item.tooltip = `@Folder("${node.folder}")`;
                item.description = node.moduleCount === 1 ? '1 module' : `${node.moduleCount} modules`;
                break;

            case 'module':
                item.iconPath = new vscode.ThemeIcon(moduleThemeIconName(node.moduleType));
                item.description = node.moduleType;
                item.contextValue = `module-${node.moduleType ?? 'standard'}`;
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

            case 'designer':
                // The vbide's arrangement: the design face announces itself as
                // the form's first child, above the handlers.
                item.iconPath = new vscode.ThemeIcon('symbol-color');
                item.contextValue = 'designer';
                item.tooltip = `Open the designer for ${node.moduleName}`;
                item.command = {
                    command: 'xlide.previewForm',
                    title: 'Open Designer',
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
            return this._getProjectFiles();
        }
        if (node.kind === 'project') {
            const modules = await this._getModules(node.filePath);
            if (this._view !== 'folders' || modules.some((m) => m.kind === 'loadError')) {
                return modules;
            }
            const tree = this._folderTreeOf(node.filePath, modules);
            return [...this._folderNodesOf(node.filePath, tree.folders), ...tree.modules];
        }
        if (node.kind === 'folder') {
            const folder = this._folderIn(node.filePath, node.folder ?? '');
            if (!folder) {
                return [];
            }
            return [...this._folderNodesOf(node.filePath, folder.folders), ...folder.modules];
        }
        if (node.kind === 'module') {
            return this._getSubs(node.filePath, node.moduleName!, node.moduleType);
        }
        return [];
    }

    /**
     * The project's folder layout, built once per module listing. Derived
     * state: refresh() clears it along with the listing it came from.
     */
    private _folderTreeOf(filePath: string, modules: XlideNode[]): FolderTree<XlideNode> {
        const key = projectNodeKey(filePath);
        let tree = this._folderTrees.get(key);
        if (!tree) {
            tree = buildFolderTree(modules);
            this._folderTrees.set(key, tree);
            this._registerFolderNodes(filePath, tree.folders);
        }
        return tree;
    }

    /**
     * A node for every folder in the project, not just the level being drawn.
     * reveal() walks a module up through folders whose parents may never have
     * been expanded, and one missing node breaks the whole walk.
     *
     * A folder that already has a node keeps it, so its expansion survives a
     * module moving in or out; only what the layout derives is refreshed.
     */
    private _registerFolderNodes(
        filePath: string,
        folders: readonly FolderTreeFolder<XlideNode>[],
    ): void {
        for (const folder of folders) {
            const key = folderNodeKey(filePath, folder.path);
            const existing = this._folderNodes.get(key);
            if (!existing) {
                this._folderNodes.set(key, {
                    kind: 'folder',
                    label: folder.name,
                    filePath,
                    folder: folder.path,
                    moduleCount: folder.moduleCount,
                });
            } else if (existing.label !== folder.name || existing.moduleCount !== folder.moduleCount) {
                existing.label = folder.name;
                existing.moduleCount = folder.moduleCount;
                // A new render id, or VS Code redraws the row it already has
                // and the count stays at what it was.
                this._bumpFolderVersions([key]);
            }
            this._registerFolderNodes(filePath, folder.folders);
        }
    }

    /**
     * Move a module because its `@Folder` annotation changed in the editor.
     * Held apart from the module listing, which is read from the container and
     * thrown away by every refresh(): what an open editor says outlives that,
     * and waiting for a save would leave the module in the folder its text has
     * just taken it out of.
     */
    setModuleFolder(filePath: string, moduleName: string, folder: string | undefined): void {
        const key = moduleNodeKey(filePath, moduleName);
        const next = folder || undefined;
        this._editorFolders.set(key, next);
        const node = this._moduleNodes.get(key);
        if (node?.folder === next) {
            // Either the tree already draws it there, or it has not listed the
            // module yet and will pick this up when it does.
            return;
        }
        if (node) { node.folder = next; }
        // The layout is derived from the listing: the folder the module left
        // may now be empty and the one it joined may be new, so it is rebuilt
        // rather than patched.
        this._folderTrees.delete(projectNodeKey(filePath));
        if (this._view === 'folders') {
            this._emitter.fire();
        }
    }

    /**
     * The module's editor closed, so the container is the truth again - the
     * edit was either saved into it or thrown away. The listing was read
     * before either happened, so it is re-read rather than trusted.
     */
    forgetModuleFolder(filePath: string, moduleName: string): void {
        if (!this._editorFolders.delete(moduleNodeKey(filePath, moduleName))) {
            return;
        }
        // A listing already in flight was read before the editor closed, so it
        // must not land in the cache this is clearing - the same reason
        // refresh() bumps the generation.
        this._generation++;
        const projectKey = projectNodeKey(filePath);
        this._modulesListCache.delete(projectKey);
        this._modulesListLoads.delete(projectKey);
        this._folderTrees.delete(projectKey);
        this._emitter.fire();
    }

    /** One folder in a project's layout, found by its dotted path. */
    private _folderIn(filePath: string, wanted: string): FolderTree<XlideNode> | undefined {
        let level: FolderTree<XlideNode> | undefined = this._folderTrees.get(projectNodeKey(filePath));
        for (const step of folderPathChain(wanted)) {
            level = level?.folders.find((folder) => folder.path === step);
        }
        return level;
    }

    /** The stable nodes for a level of folders, in the order they are drawn. */
    private _folderNodesOf(
        filePath: string,
        folders: readonly FolderTreeFolder<XlideNode>[],
    ): XlideNode[] {
        return folders.map((folder) => this._folderNodes.get(folderNodeKey(filePath, folder.path))!);
    }

    private async _getProjectFiles(): Promise<XlideNode[]> {
        if (this._projectFilesCache) {
            return this._projectFilesCache;
        }
        if (this._projectFilesLoad) {
            return this._projectFilesLoad;
        }
        const load = this._loadProjectFiles();
        this._projectFilesLoad = load;
        try {
            const nodes = await load;
            this._projectFilesCache = nodes;
            return nodes;
        } finally {
            if (this._projectFilesLoad === load) {
                this._projectFilesLoad = undefined;
            }
        }
    }

    private async _loadProjectFiles(): Promise<XlideNode[]> {
        const uris = await findMacroContainerFiles();
        return uris
            .map((uri) => {
                let node = this._projectNodes.get(uri.fsPath);
                if (!node) {
                    node = { kind: 'project', label: fileNameForDisplay(uri.fsPath), filePath: uri.fsPath };
                    this._projectNodes.set(uri.fsPath, node);
                }
                return node;
            });
    }

    private async _getModules(filePath: string): Promise<XlideNode[]> {
        this._cancelProtectionTimer(filePath);
        try {
            const cacheKey = projectNodeKey(filePath);
            let modules = this._modulesListCache.get(cacheKey);
            let loadGeneration: number | undefined;
            if (!modules) {
                let load = this._modulesListLoads.get(cacheKey);
                if (!load) {
                    load = this._bridge.call<Array<{ name: string; type: string; filePath?: string; folder?: string }>>(
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
                    // An open editor's annotation outranks the container's,
                    // and is re-applied on every listing so a refresh does not
                    // snap the module back to where it was saved.
                    const folder = this._editorFolders.has(key)
                        ? this._editorFolders.get(key)
                        : (m.folder || undefined);
                    let node = this._moduleNodes.get(key);
                    if (!node) {
                        node = {
                            kind: 'module',
                            label: m.name,
                            filePath,
                            moduleName: m.name,
                            moduleType: m.type,
                            ...(m.filePath ? { moduleFilePath: m.filePath } : {}),
                            ...(folder ? { folder } : {}),
                        };
                        if (populateNodeMap) {
                            this._moduleNodes.set(key, node);
                        }
                    } else {
                        node.folder = folder;
                    }
                    return node;
                });
            this._scheduleProtectionLoad(filePath);
            return nodes;
        } catch (err) {
            vscode.window.showErrorMessage(`XLIDE: Failed to list modules in "${fileNameForDisplay(filePath)}": ${err}`);
            // Return a retry placeholder, never [] - VS Code caches resolved
            // children, so an empty result would leave the project permanently
            // empty after a transient failure (e.g. Excel holding the file).
            return [this._loadErrorNode(filePath, undefined, err)];
        }
    }

    /**
     * Lazily fetch the project's VBA protection/signature state and, once
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
                const node = this._projectNodes.get(filePath);
                if (node) {
                    node.isPasswordProtected = info.isPasswordProtected;
                    node.isSigned = info.isSigned;
                    this._emitter.fire(node);
                }
            } catch (err) {
                // Badge is best-effort; log the probe failure without surfacing it.
                this._out?.appendLine(`[projectExplorer] Protection probe failed for "${fileNameForDisplay(filePath)}": ${err}`);
            } finally {
                this._protectionLoads.delete(filePath);
            }
        })();
        this._protectionLoads.set(filePath, load);
        await load;
    }

    private async _getSubs(filePath: string, moduleName: string, moduleType?: string): Promise<XlideNode[]> {
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
                if (this._generation !== generation) {
                    // The rows built below are as stale as the list they come
                    // from, so they are handed back for this render and not
                    // kept: keeping them would win over the re-fetch that the
                    // cleared list cache is about to make, permanently.
                    this._scheduleProtectionLoad(filePath);
                    return this._buildSubNodes(subs, filePath, moduleName, moduleType);
                }
                this._subsListCache.set(cacheKey, subs);
            }
            // Built once and kept: treeView.reveal() matches the element it is
            // given against the ones the tree drew, so a fresh object per
            // render would leave the caret's own procedure unfindable.
            let nodes = this._subNodes.get(cacheKey);
            if (!nodes) {
                nodes = this._buildSubNodes(subs, filePath, moduleName, moduleType);
                this._subNodes.set(cacheKey, nodes);
            }
            this._scheduleProtectionLoad(filePath);
            return nodes;
        } catch (err) {
            vscode.window.showErrorMessage(`XLIDE: Failed to list procedures in "${moduleName}" (${fileNameForDisplay(filePath)}): ${err}`);
            return [this._loadErrorNode(filePath, moduleName, err)];
        }
    }

    private _buildSubNodes(
        subs: ReadonlyArray<{ name: string; kind: string; line: number }>,
        filePath: string,
        moduleName: string,
        moduleType: string | undefined,
    ): XlideNode[] {
        const nodes: XlideNode[] = subs.map((s) => ({
            kind: 'sub' as const,
            label: `${s.kind} ${s.name}`,
            filePath,
            moduleName,
            line: s.line,
        }));
        // A VB6 form, UserControl, or PropertyPage opens in the designer too,
        // drawn from its own header (roadmap_vb6_support.md, Slice 5).
        const vb6Designer = isVb6ProjectPath(filePath)
            && (moduleType === 'usercontrol' || moduleType === 'propertypage');
        // An Access form or report is a design first: it has one whether or
        // not Access has ever opened its code window.
        const accessDesign = moduleType === 'accessform' || moduleType === 'accessreport';
        if (moduleType === 'userform' || vb6Designer || accessDesign) {
            // The designer sits FIRST under its form, above the handlers - the
            // xlide vbide arrangement: the design comes before the code that
            // answers it, and a fixed position means the row never moves as
            // procedures are added and renamed.
            nodes.unshift({ kind: 'designer', label: 'Designer', filePath, moduleName });
        }
        return nodes;
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

    private _refreshNonActiveProjectExpansions(activeProjectKey: string | undefined): void {
        for (const node of this._projectNodes.values()) {
            const key = projectNodeKey(node.filePath);
            if (key === activeProjectKey) {
                continue;
            }
            this._projectRenderVersions.set(key, (this._projectRenderVersions.get(key) ?? 0) + 1);
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

function projectNodeKey(filePath: string): string {
    return projectIdentityKey(filePath);
}

function folderNodeKey(filePath: string, folder: string): string {
    return `${projectNodeKey(filePath)}::${folder.toLowerCase()}`;
}

/** The folder one level out, or '' for a folder already at the project's root. */
function parentFolderPath(folder: string): string {
    const cut = folder.lastIndexOf('.');
    return cut === -1 ? '' : folder.slice(0, cut);
}

function sameOrder(left: readonly string[], right: readonly string[]): boolean {
    return left.length === right.length && left.every((value, i) => value === right[i]);
}
