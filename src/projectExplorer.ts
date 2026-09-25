import * as vscode from 'vscode';
import * as path from 'path';
import { ProjectEngine } from './projectEngine';
import { moduleIdentityKey, projectIdentityKey } from './xlideFileSystem';
import { compareVbaModulesForTreeOrder, moduleThemeIconName } from './moduleDisplay';
import { buildFolderTree, folderPathChain, type FolderTree, type FolderTreeFolder } from './folderTree';
import type { XlideExplorerView } from './globalSettings';
import { canAddVbaProjectTo, containerAppNameForPath, containerContextValue, isVb6ProjectPath } from './macroContainerUi';
import { findMacroContainerFiles } from './macroContainerDiscovery';
import { hasPendingAgentReview, pendingAgentReviewModules } from './xlideAgentDiff';
import type { GitChangeMarksSource } from './gitChangeMarks';
import {
    AGENT_REVIEW_COLOR_ID,
    moduleDecorationUri,
    projectDecorationUri,
} from './agentReviewDecorations';
import { startPerformanceTrace } from './performanceTrace';
import { osPlatform } from './util/osPlatform';
import { ShapeRows, type ShapeRowContext } from './shapeRows';
import type { ShapeInfo } from './vba/shapes';

export type XlideNodeKind = 'project' | 'folder' | 'module' | 'designer' | 'sub' | 'loadError' | 'empty'
    // The shape rows, drawn by ShapeRows (shapeRows.ts): a Shapes or Slides
    // folder, a slide or story, a shape, and the row an empty one shows.
    | 'shapes' | 'surface' | 'shape' | 'noShapes';

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
     * sub only: which of the module's rows with this label it is, from 0.
     * Two share a label only while a duplicate procedure is being edited.
     */
    occurrence?: number;
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
    /**
     * empty only: whether the file holds a VBA project with nothing in it, as
     * opposed to holding none at all. Both list no modules; only the first can
     * take a new one, so the row says which it is.
     */
    hasVbaProject?: boolean;
    /** Workbook only: VBA project carries a digital signature. */
    isSigned?: boolean;
    /** module only: what a document module stands for (worksheet, workbook, chart, document). */
    documentType?: string;
    /** module only: the sheet the module stands for, when the workbook lists it; the row is named for it. */
    sheetName?: string;
    /**
     * shapes only: a module's Shapes folder, a sheet row's Shapes folder
     * (surface), a presentation's Slides, a workbook's Sheets, or the folder
     * of its sheets with no module and no shapes (bareSheets).
     */
    shapeFolder?: 'module' | 'surface' | 'slides' | 'sheets' | 'bareSheets';
    /** surface and shape only: the sheet, slide or Word story, as the shape tools name it. */
    surface?: string;
    /** shape only: the shape as the file lists it. */
    shape?: ShapeInfo;
    /** shape only: its name, after the names of the groups it is in. */
    shapePath?: string[];
    /** surface and the sheets folder: how many rows are under it. */
    itemCount?: number;
}

export class ProjectExplorer implements vscode.TreeDataProvider<XlideNode>, vscode.Disposable {
    private _emitter = new vscode.EventEmitter<XlideNode | undefined | null | void>();
    readonly onDidChangeTreeData = this._emitter.event;
    private readonly _rowsReplaced = new vscode.EventEmitter<{ filePath: string; moduleName: string } | undefined>();
    /**
     * Rows were drawn again from scratch, so a selection or a reveal made on
     * the old ones may be gone: one module's rows below it, or (undefined)
     * the whole tree's.
     */
    readonly onDidReplaceRows = this._rowsReplaced.event;


    // Stable node references required by treeView.reveal()
    // Every cache below is keyed by identity, never by a raw path: one file
    // reaches these under more than one spelling (uri.fsPath vs a decoded
    // xlide-vba:// path), and keying raw meant the same workbook landed in
    // two entries.
    private _projectNodes = new Map<string, XlideNode>(); // key: projectNodeKey
    private _moduleNodes = new Map<string, XlideNode>(); // key: moduleNodeKey
    private _projectRenderVersions = new Map<string, number>();
    private _moduleRenderVersions = new Map<string, number>();
    private _projectFilesCache: XlideNode[] | undefined;
    private _projectFilesLoad: Promise<XlideNode[]> | undefined;
    // listModules cache: avoids repeated bridge round-trips while the tree is
    // expanded.  Cleared on refresh() so edits always re-fetch.
    private _modulesListCache = new Map<string, ModuleListing[]>();
    private _modulesListLoads = new Map<string, Promise<ModuleListing[]>>();
    // Bumped on every refresh(). An in-flight load captured before a refresh must
    // not write its now-stale result into the freshly-cleared cache (which would
    // leave a just-added module invisible until the next refresh).
    private _generation = 0;
    private _subsListCache = new Map<string, Array<{ name: string; kind: string; line: number }>>();
    private _subsListLoads = new Map<string, Promise<Array<{ name: string; kind: string; line: number }>>>();
    // The drawn sub/designer rows of a module, kept so reveal() can name one.
    private _subNodes = new Map<string, XlideNode[]>();
    // Protection-state cache: {isPasswordProtected, isSigned} per projectNodeKey.
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
    private _folderNodes = new Map<string, XlideNode>(); // key: folderNodeKey
    private _folderRenderVersions = new Map<string, number>();
    private _folderTrees = new Map<string, FolderTree<XlideNode>>();
    // Folders the editor's own module opened, and the ones the user opened or
    // shut by hand, which outrank it until the attention genuinely moves.
    private _openFolderKeys = new Set<string>();
    private _manualFolderStates = new Map<string, boolean>();
    // The folders VS Code shows open, by whatever hand: a reveal's, the
    // user's, the editor's. A fold has to reach every one of them, not only
    // the ones the two sets above remember, or a folder a reveal opened after
    // they let go of it keeps its state through every redraw.
    private _expandedFolderKeys = new Set<string>();
    private _activeFolderChain: string[] = [];
    // What an open editor says a module's @Folder annotation is, which outranks
    // the listing until that editor closes. Survives refresh() on purpose.
    private _editorFolders = new Map<string, string | undefined>();

    // The view drawing this tree, for the reveal a row click asks for.
    private _treeView: vscode.TreeView<XlideNode> | undefined;
    // While a probe runs: the rows whose children VS Code asked for.
    private _childrenProbe: string[] | undefined;
    // The shape rows under worksheet and document modules and presentations.
    private readonly _shapes: ShapeRows;

    constructor(
        private readonly _bridge: ProjectEngine,
        private readonly _out?: vscode.OutputChannel,
        private readonly _gitMarks?: GitChangeMarksSource,
    ) {
        this._shapes = new ShapeRows(_bridge, (node) => this._emitter.fire(node), _out);
    }

    dispose(): void {
        this._clearProtectionTimers();
        this._emitter.dispose();
        this._rowsReplaced.dispose();
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

    /**
     * Read everything again. The render ids are kept: VS Code keeps the state
     * of a row whose id it already holds and applies the item's own state only
     * to an id it has not seen. Resetting them to 0 handed rows ids VS Code
     * could still hold from before, in whatever state they had then, so a
     * project the editor had just moved into came back folded.
     */
    refresh(): void {
        this._generation++;
        this._projectNodes.clear();
        this._moduleNodes.clear();
        this._folderNodes.clear();
        this._folderTrees.clear();
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
        this._shapes.clear();
        this._emitter.fire();
        this._rowsReplaced.fire(undefined);
    }

    /**
     * A file changed, by a shape edit or a save from outside: its shapes are
     * read again when next shown, and the shape rows someone opened are
     * redrawn. A workbook's Sheets folder is redrawn when its sheets changed,
     * and after XLIDE's own shape edit (shapesChanged), since a sheet with
     * its first shape, or without its last, moves within it. Nothing else in
     * the tree moves.
     */
    refreshShapes(filePath: string, options: { shapesChanged?: boolean } = {}): void {
        this._shapes.refresh(filePath, options);
    }

    /** The surface and shape a shape row, surface row or Shapes folder stands for. */
    shapeContextOf(node: XlideNode): ShapeRowContext | undefined {
        return this._shapes.contextOf(node);
    }

    /**
     * The surface a Shapes folder or surface row adds a shape to, reading the
     * file when the folder has never been opened.
     */
    shapeSurfaceOf(node: XlideNode): Promise<ShapeRowContext | undefined> {
        return this._shapes.surfaceOf(node);
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
            this._rowsReplaced.fire({ filePath, moduleName });
        }
    }

    /**
     * Redraw the rows that show whether an agent edit awaits review: the
     * module, and in the folder layout every folder it sits under. The project
     * row needs no redraw; its decoration follows the pending set by itself.
     *
     * Rows are fired without a new render id, so a folder the user has open
     * stays open.
     */
    refreshAgentReviewMarks(filePath: string, moduleName: string): void {
        const module = this._findModuleNode(filePath, moduleName);
        this.refreshModuleSubs(module?.filePath ?? filePath, module?.moduleName ?? moduleName);
        if (this._view !== 'folders' || !module?.folder) {
            return;
        }
        for (const step of folderPathChain(module.folder)) {
            const folder = this._folderNodes.get(folderNodeKey(module.filePath, step));
            if (folder) {
                this._emitter.fire(folder);
            }
        }
    }

    /**
     * Redraw every row of a project that carries a git mark: the modules,
     * whose `M`/`A` badges follow the marks, and the project row's count.
     * Fired without a new render id, so nothing the user has open closes.
     */
    refreshGitMarks(filePath: string): void {
        const project = projectIdentityKey(filePath);
        for (const node of this._moduleNodes.values()) {
            if (projectIdentityKey(node.filePath) === project) {
                this._emitter.fire(node);
            }
        }
        const projectNode = this._projectNodes.get(project);
        if (projectNode) {
            this._emitter.fire(projectNode);
        }
    }

    /**
     * The loaded module row for a module, however the caller spelled the path
     * or the name. An agent tool and the tree can name the same module in
     * different cases, and the row cache is keyed on the tree's spelling.
     */
    private _findModuleNode(filePath: string, moduleName: string): XlideNode | undefined {
        const exact = this._moduleNodes.get(moduleNodeKey(filePath, moduleName));
        if (exact) {
            return exact;
        }
        const project = projectIdentityKey(filePath);
        const wanted = moduleName.toLowerCase();
        for (const node of this._moduleNodes.values()) {
            if (projectIdentityKey(node.filePath) === project && node.moduleName?.toLowerCase() === wanted) {
                return node;
            }
        }
        return undefined;
    }

    /** Whether any module under this folder, at any depth, awaits review. */
    private _folderHasPendingAgentReview(folder: XlideNode): boolean {
        if (pendingAgentReviewModules(folder.filePath).length === 0) {
            return false;
        }
        const project = projectIdentityKey(folder.filePath);
        const wanted = folderNodeKey(folder.filePath, folder.folder ?? '');
        for (const module of this._moduleNodes.values()) {
            if (projectIdentityKey(module.filePath) !== project || !module.folder) {
                continue;
            }
            if (!hasPendingAgentReview(module.filePath, module.moduleName ?? '')) {
                continue;
            }
            if (folderPathChain(module.folder).some((step) => folderNodeKey(module.filePath, step) === wanted)) {
                return true;
            }
        }
        return false;
    }

    /** Required by treeView.reveal() - walks xlsm -> (folder) -> module -> sub. */
    getParent(node: XlideNode): XlideNode | undefined {
        if (node.kind === 'folder') {
            const parent = parentFolderPath(node.folder ?? '');
            return parent
                ? this._folderNodes.get(folderNodeKey(node.filePath, parent))
                : this._projectNodes.get(projectNodeKey(node.filePath));
        }
        if (node.kind === 'module') {
            // A sheet's module sits under the workbook's Sheets folder,
            // whatever layout the other modules are in.
            const sheets = this._shapes.parentOf(node);
            if (sheets) {
                return sheets;
            }
            if (this._view === 'folders' && node.folder) {
                return this._folderNodes.get(folderNodeKey(node.filePath, node.folder));
            }
            return this._projectNodes.get(projectNodeKey(node.filePath));
        }
        if (node.kind === 'sub') {
            return this._moduleNodes.get(moduleNodeKey(node.filePath, node.moduleName ?? ''));
        }
        if (node.kind === 'loadError' || node.kind === 'empty') {
            return node.moduleName
                ? this._moduleNodes.get(moduleNodeKey(node.filePath, node.moduleName))
                : this._projectNodes.get(projectNodeKey(node.filePath));
        }
        if (isShapeRow(node)) {
            return this._shapes.parentOf(node);
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
        const project = this._projectNodes.get(projectNodeKey(node.filePath));
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

    /**
     * The row a project with no code gets. It is NOT a load failure: a
     * workbook saved as .xlsm before the first macro is written carries no
     * VBA project at all, and XLIDE read it perfectly well. Offering "click
     * to retry" there sent people looking for a problem that was not there.
     */
    private _emptyNode(filePath: string, hasVbaProject: boolean): XlideNode {
        return {
            kind: 'empty',
            label: hasVbaProject ? 'No modules yet' : 'No VBA in this file yet',
            filePath,
            hasVbaProject,
        };
    }

    /**
     * Whether the tree lists the file as a project, from the listing it draws.
     * That listing is kept until the next refresh, so asking is cheap.
     */
    async listsProject(filePath: string): Promise<boolean> {
        const key = projectNodeKey(filePath);
        return (await this._getProjectFiles()).some((node) => projectNodeKey(node.filePath) === key);
    }

    /** Returns the cached module node, if the tree has loaded it. */
    getModuleNode(filePath: string, moduleName: string): XlideNode | undefined {
        return this._moduleNodes.get(moduleNodeKey(filePath, moduleName));
    }

    /**
     * A module's row, loading what the tree has not drawn yet to reach it:
     * the list of projects, and the project's modules and, in the folder
     * layout, its folders. reveal() can only walk up through rows that exist,
     * and a project nobody expanded has none. Undefined when no project in
     * the workspace holds the module.
     */
    async resolveModuleNode(filePath: string, moduleName: string): Promise<XlideNode | undefined> {
        const cached = this.getModuleNode(filePath, moduleName);
        if (cached) {
            return cached;
        }
        const project = (await this._getProjectFiles())
            .find((node) => projectNodeKey(node.filePath) === projectNodeKey(filePath));
        if (!project) {
            return undefined;
        }
        await this._getChildren(project);
        return this.getModuleNode(filePath, moduleName);
    }

    /** A procedure's row, listing the module's procedures if the tree has not. */
    async resolveProcedureNode(filePath: string, moduleName: string, label: string): Promise<XlideNode | undefined> {
        const module = await this.resolveModuleNode(filePath, moduleName);
        if (!module) {
            return undefined;
        }
        if (!this._subNodes.has(moduleNodeKey(filePath, moduleName))) {
            await this._getChildren(module);
        }
        return this.getProcedureNode(filePath, moduleName, label);
    }

    /**
     * A row's identity: the same for every node object that draws it, and
     * unlike its TreeItem id, the same across the re-renders that fold it.
     */
    rowIdentity(node: XlideNode): string {
        switch (node.kind) {
            case 'project':
                return `project::${projectNodeKey(node.filePath)}`;
            case 'folder':
                return `folder::${folderNodeKey(node.filePath, node.folder ?? '')}`;
            case 'module':
                return `module::${moduleNodeKey(node.filePath, node.moduleName ?? '')}`;
            case 'shapes':
            case 'surface':
            case 'shape':
            case 'noShapes':
                // Its tree item id, which names the file, the surface and the
                // shape's place among groups.
                return `${node.kind}::${this._shapes.treeItem(node).id}`;
            default:
                return `${node.kind}::${moduleNodeKey(node.filePath, node.moduleName ?? '')}::${node.label}`;
        }
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

    /** The view that draws this tree; set once, right after it is created. */
    attachTreeView(treeView: vscode.TreeView<XlideNode>): void {
        this._treeView = treeView;
    }

    /**
     * For tests and development only: what the view shows. VS Code has no
     * call that says which rows are expanded, but a full redraw re-reads the
     * children of the expanded rows and no others, so this redraws and notes
     * whose children are asked for, until the asking stops.
     */
    async probeViewState(): Promise<{ expanded: string[]; selected: string[]; activeModule: string | undefined }> {
        const asked: string[] = [];
        this._childrenProbe = asked;
        this._emitter.fire();
        try {
            let seen = -1;
            for (let quiet = 0; quiet < 4;) {
                await new Promise((resolve) => setTimeout(resolve, 100));
                quiet = asked.length === seen ? quiet + 1 : 0;
                seen = asked.length;
            }
        } finally {
            this._childrenProbe = undefined;
        }
        const active = this._activeModuleKey === undefined ? undefined : this._moduleNodes.get(this._activeModuleKey);
        return {
            expanded: [...new Set(asked)].sort(),
            selected: (this._treeView?.selection ?? []).map(describeNode),
            activeModule: active ? describeNode(active) : undefined,
        };
    }

    /**
     * A module row was clicked: it becomes the module the tree keeps open, and
     * opens if it was folded. Following the editor does this for any other
     * module, but a click on the module already in front changes no editor,
     * so a row folded by hand stayed folded.
     */
    async expandModuleRow(node: XlideNode): Promise<void> {
        if (node.kind !== 'module' || !node.moduleName || !this._treeView) {
            return;
        }
        this.setActiveModule(node.filePath, node.moduleName);
        try {
            await this._treeView.reveal(node, { select: true, focus: false, expand: true });
        } catch {
            // The row went with a refresh; the next render draws it open.
        }
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
     * Folds a module unless it is the one the accordion keeps open. A reveal
     * still under way when the editor left the module opens it after the
     * accordion has moved on, and nothing else would fold it again.
     */
    foldModuleUnlessActive(filePath: string, moduleName: string): void {
        const key = moduleNodeKey(filePath, moduleName);
        if (key !== this._activeModuleKey) {
            this._refreshModuleExpansion(key);
        }
    }

    /** Collapse a module whose last editor tab closed, without folding its project. */
    clearActiveModule(filePath: string, moduleName: string): void {
        const key = moduleNodeKey(filePath, moduleName);
        if (this._activeModuleKey !== key) { return; }
        this._activeModuleKey = undefined;
        this._refreshModuleExpansion(key);
        this.collapseAllFolders();
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
            ...this._expandedFolderKeys,
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
     * What VS Code reports for a folder row, whoever opened or shut it. A
     * reveal under way when the tree let go of its folders (a tab switch
     * leaves no editor active for a moment, which folds everything) opens
     * them again in the view alone; noted here, the next move folds them.
     */
    noteFolderExpanded(node: XlideNode, expanded: boolean): void {
        if (node.kind !== 'folder') { return; }
        const key = folderNodeKey(node.filePath, node.folder ?? '');
        if (expanded) {
            this._expandedFolderKeys.add(key);
        } else {
            this._expandedFolderKeys.delete(key);
        }
    }

    /**
     * Fold every folder, in every project. The last editor closing leaves no
     * module for the tree to follow, so it goes back to its resting shape.
     */
    collapseAllFolders(): void {
        if (this._openFolderKeys.size === 0 && this._manualFolderStates.size === 0 && this._expandedFolderKeys.size === 0) { return; }
        const touched = new Set([...this._openFolderKeys, ...this._manualFolderStates.keys(), ...this._expandedFolderKeys]);
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
        if (isShapeRow(node)) {
            return this._shapes.treeItem(node);
        }
        const isActiveModule =
            node.kind === 'module' &&
            moduleNodeKey(node.filePath, node.moduleName ?? '') === this._activeModuleKey;
        const isActiveProject =
            node.kind === 'project' &&
            projectNodeKey(node.filePath) === this._activeProjectKey;
        const isOpenFolder = node.kind === 'folder' && this._folderIsOpen(node);
        const item = new vscode.TreeItem(
            node.label,
            node.kind === 'sub' || node.kind === 'designer'
                || node.kind === 'loadError' || node.kind === 'empty'
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
            // Not the line, which every edit above the procedure moves: a new
            // id is a new row to VS Code, and the selection on the old one,
            // the caret's procedure, went with it on each save.
            item.id = `s::${node.filePath}::${node.moduleName}::${node.label}::${node.occurrence ?? 0}`;
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
                // Always carried, so the row can light up the moment an agent
                // edit lands in it without redrawing a tree the user has open.
                item.resourceUri = projectDecorationUri(node.filePath);
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

            case 'folder': {
                item.contextValue = 'folder';
                // The annotation that makes the folder, so a nested one reads
                // as the whole path rather than just its last segment.
                item.tooltip = `@Folder("${node.folder}")`;
                const count = node.moduleCount === 1 ? '1 module' : `${node.moduleCount} modules`;
                // A folder row gets no decoration URI: the `folder` theme icon
                // switches to the file-icon theme once a row has one. The
                // pending mark goes on the icon and the description instead.
                if (this._folderHasPendingAgentReview(node)) {
                    item.iconPath = new vscode.ThemeIcon('folder', new vscode.ThemeColor(AGENT_REVIEW_COLOR_ID));
                    item.description = `${count} ● agent edit`;
                } else {
                    item.iconPath = new vscode.ThemeIcon('folder');
                    item.description = count;
                }
                break;
            }

            case 'module':
                item.iconPath = new vscode.ThemeIcon(moduleThemeIconName(node.moduleType));
                item.description = node.moduleType;
                item.contextValue = `module-${node.moduleType ?? 'standard'}`;
                if (hasPendingAgentReview(node.filePath, node.moduleName ?? '')) {
                    // An agent wrote this module and nobody has kept or
                    // reverted it yet. The decoration colours the name and
                    // badges it; the tinted icon and the description say the
                    // same thing to a theme that ignores decoration colours.
                    item.description = `${node.moduleType} ● agent edit`;
                    item.contextValue += '-agent-pending';
                    item.iconPath = new vscode.ThemeIcon(
                        moduleThemeIconName(node.moduleType),
                        new vscode.ThemeColor(AGENT_REVIEW_COLOR_ID),
                    );
                    item.resourceUri = moduleDecorationUri(node.filePath, node.moduleName ?? '');
                    // Set, because a row with a resourceUri and no tooltip
                    // shows the URI on hover.
                    item.tooltip = `${node.label}\n\nAn agent edited this module. `
                        + 'Review, keep or revert the change with the buttons on this row.';
                } else {
                    // Changed since the last commit: the decoration badges the
                    // row `M` or `A` in the Explorer's colours. Asking for the
                    // marks is what schedules them the first time.
                    const change = this._gitMarks?.marksFor(node.filePath)
                        ?.byModule.get((node.moduleName ?? '').toLowerCase());
                    if (change) {
                        item.resourceUri = moduleDecorationUri(node.filePath, node.moduleName ?? '');
                        item.tooltip = `${node.label}\n\n${change === 'modified'
                            ? 'Modified since the last commit.'
                            : 'Not in the last commit.'} Compare it with Git HEAD from this row's menu.`;
                    }
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

            case 'empty':
                // Nothing went wrong here, so nothing warns and nothing
                // offers a retry: the icon and the wording say the file is
                // empty, which is a state it is allowed to be in.
                item.iconPath = new vscode.ThemeIcon('info');
                item.contextValue = node.hasVbaProject ? 'emptyProject'
                    : canAddVbaProjectTo(node.filePath) ? 'noVbaProject' : 'noVbaProjectFixed';
                if (node.hasVbaProject) {
                    item.tooltip = 'XLIDE read this file without trouble. Its VBA project has no modules in it'
                        + " yet; add one from the file's row above.";
                } else if (osPlatform() === 'web' || !canAddVbaProjectTo(node.filePath)) {
                    // The browser cannot read the templates a project starts
                    // from, and a legacy file takes none from XLIDE.
                    item.tooltip = `XLIDE read this file without trouble. It has no VBA project in it at all,`
                        + ` which is how ${containerAppNameForPath(node.filePath)} saves a`
                        + ' macro-enabled file that has never held a macro. Write one macro in'
                        + ` ${containerAppNameForPath(node.filePath)} and save: the project appears here,`
                        + ' and everything after that is XLIDE\'s.';
                } else {
                    item.tooltip = `XLIDE read this file without trouble. It has no VBA project in it at all,`
                        + ` which is how ${containerAppNameForPath(node.filePath)} saves a`
                        + ' macro-enabled file that has never held a macro. Click to add one.';
                    item.command = {
                        command: 'xlide.addVbaProject',
                        title: 'Add VBA Project',
                        arguments: [node],
                    };
                }
                break;
        }

        return item;
    }

    async getChildren(node?: XlideNode): Promise<XlideNode[]> {
        if (node && this._childrenProbe) {
            this._childrenProbe.push(describeNode(node));
        }
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
            const failed = modules.some((m) => m.kind === 'loadError');
            // A presentation's slides, and a workbook's sheets, lead the
            // project's rows, and a sheet's module is drawn under Sheets; a
            // listing that failed shows only that.
            const rows = failed ? { folders: [], modules } : await this._shapes.projectRows(node, modules);
            if (modules.length === 0) {
                return [...rows.folders, this._emptyNode(node.filePath, await this._hasVbaProject(node.filePath))];
            }
            if (this._view !== 'folders' || failed) {
                return [...rows.folders, ...rows.modules];
            }
            const tree = this._folderTreeOf(node.filePath, rows.modules);
            return [...rows.folders, ...this._folderNodesOf(node.filePath, tree.folders), ...tree.modules];
        }
        if (node.kind === 'folder') {
            const folder = this._folderIn(node.filePath, node.folder ?? '');
            if (!folder) {
                return [];
            }
            return [...this._folderNodesOf(node.filePath, folder.folders), ...folder.modules];
        }
        if (node.kind === 'module') {
            const subs = await this._getSubs(node.filePath, node.moduleName!, node.moduleType);
            // A worksheet's or Word document's shapes, above its procedures,
            // where a form's Designer row sits.
            const shapes = await this._shapes.moduleFolder(node);
            return shapes ? [shapes, ...subs] : subs;
        }
        if (isShapeRow(node)) {
            return this._shapes.children(node, async () => (await this._getModules(node.filePath))
                .filter((m) => m.kind === 'module'));
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
            // The module being edited takes the open folders with it, or its
            // row would sit in a folder nobody opened.
            if (key === this._activeModuleKey) {
                this._followFoldersTo(filePath, moduleName);
            }
            this._emitter.fire();
            this._rowsReplaced.fire({ filePath, moduleName });
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
                const projectKey = projectNodeKey(uri.fsPath);
                let node = this._projectNodes.get(projectKey);
                if (!node) {
                    node = { kind: 'project', label: fileNameForDisplay(uri.fsPath), filePath: uri.fsPath };
                    this._projectNodes.set(projectKey, node);
                }
                return node;
            });
    }

    /**
     * Whether the file holds a VBA project at all, asked only when it listed
     * no modules - so it decides the wording of one row and nothing else.
     * Cheap: the container is already parsed and cached by the listing that
     * just ran, and this reads one part name out of it.
     *
     * A listing that succeeded makes this succeed too; if it somehow does
     * not, the row says the neutral thing rather than claiming the file is
     * empty of VBA.
     */
    private async _hasVbaProject(filePath: string): Promise<boolean> {
        try {
            const info = await this._bridge.call<{ hasVbaProject: boolean }>(
                'hasVbaProject', { path: filePath },
            );
            return info.hasVbaProject;
        } catch {
            return true;
        }
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
                    load = this._bridge.call<ModuleListing[]>(
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
                            ...(m.documentType ? { documentType: m.documentType } : {}),
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
        const key = projectNodeKey(filePath);
        if (this._protectionCache.has(key)) { return; }
        const existing = this._protectionLoads.get(key);
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
                this._protectionCache.set(key, info);
                const node = this._projectNodes.get(key);
                if (node) {
                    node.isPasswordProtected = info.isPasswordProtected;
                    node.isSigned = info.isSigned;
                    this._emitter.fire(node);
                }
            } catch (err) {
                // Badge is best-effort; log the probe failure without surfacing it.
                this._out?.appendLine(`[projectExplorer] Protection probe failed for "${fileNameForDisplay(filePath)}": ${err}`);
            } finally {
                this._protectionLoads.delete(key);
            }
        })();
        this._protectionLoads.set(key, load);
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
        const seen = new Map<string, number>();
        const nodes: XlideNode[] = subs.map((s) => {
            const label = `${s.kind} ${s.name}`;
            const occurrence = seen.get(label) ?? 0;
            seen.set(label, occurrence + 1);
            return { kind: 'sub' as const, label, filePath, moduleName, line: s.line, occurrence };
        });
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
        const key = projectNodeKey(filePath);
        if (this._protectionCache.has(key) || this._protectionLoads.has(key)) {
            return;
        }
        if (this._protectionTimers.has(key)) {
            return;
        }
        const timer = setTimeout(() => {
            this._protectionTimers.delete(key);
            void this._loadProtection(filePath);
        }, PROTECTION_PROBE_IDLE_DELAY_MS);
        (timer as unknown as { unref?: () => void }).unref?.();
        this._protectionTimers.set(key, timer);
    }

    private _cancelProtectionTimer(filePath: string): void {
        const key = projectNodeKey(filePath);
        const timer = this._protectionTimers.get(key);
        if (!timer) {
            return;
        }
        clearTimeout(timer);
        this._protectionTimers.delete(key);
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

/** What listModules gives the tree for each module. */
interface ModuleListing {
    name: string;
    type: string;
    documentType?: string;
    filePath?: string;
    folder?: string;
}

/** Whether ShapeRows draws the row. */
function isShapeRow(node: XlideNode): boolean {
    return node.kind === 'shapes' || node.kind === 'surface' || node.kind === 'shape' || node.kind === 'noShapes';
}

/** A row as a probe reports it: its kind, and the name it shows. */
function describeNode(node: XlideNode): string {
    const name = node.kind === 'project' ? fileNameForDisplay(node.filePath)
        : node.kind === 'folder' ? node.folder ?? node.label
            : node.kind === 'module' ? node.moduleName ?? node.label
                : node.label;
    return `${node.kind}:${name}`;
}

function fileNameForDisplay(filePath: string): string {
    return filePath.split(/[\\/]/).filter(Boolean).pop() ?? filePath;
}

/**
 * Every node key goes through the identity helpers, the way folderNodeKey
 * does. Using the raw path and name meant two strings naming one module
 * produced two keys: in a browser the tree's path comes from uri.fsPath
 * ('\Book.xlsm') while the active module's comes from decoding its
 * xlide-vba:// URI ('/Book.xlsm'), so the row never matched
 * _activeModuleKey and rendered collapsed. The raw name was a latent bug of
 * the same shape, since VBA module names are case-insensitive.
 *
 * Only ever used as a map key and as part of a TreeItem id, never parsed
 * back or shown.
 */
function moduleNodeKey(filePath: string, moduleName: string): string {
    return `${projectNodeKey(filePath)}::${moduleIdentityKey(moduleName)}`;
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
