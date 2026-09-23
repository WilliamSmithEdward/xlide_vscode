// The explorer following the editor: the module being edited open in the
// tree with the others folded (the accordion), and the row of the procedure
// the caret is in selected, the way the VBE's own explorer marks it.
//
// Every input only asks for a pass: an editor switch, the caret crossing into
// another procedure, a tab closing, the tree shown again, the setting turned
// on. A pass reads where the editor is when it runs, loads the rows it needs,
// and reveals once. Passes run one at a time after a short quiet period, and a
// pass that a newer request overtakes stops at its next step instead of
// finishing, so under rapid switching only the last editor is revealed.
//
// Before this, each input revealed on its own. Reveals overlapped, and one
// that finished late expanded a module the editor had already left. VS Code
// reports that expansion exactly as it reports a click on the row, and a
// click makes a module the one the accordion keeps open, so the tree then
// followed the stale module instead of the editor. A module whose rows the
// tree had not loaded yet was not revealed at all, and nothing tried again.
//
// So the expansions a pass's own reveal causes are told apart from the
// user's: they are the rows on the way to what it revealed, and they stay
// attributed to it for a moment after the reveal settles, because the events
// can arrive after it.
//
// Rows drawn again can undo a reveal as well. A refresh replaces every row,
// a save the rows under its module, a changed @Folder moves a module to
// another folder; a reveal under way at that moment was resolving rows that
// are gone, and VS Code drops it without a word. Rows keep their ids across a
// refresh, so VS Code keeps their state, and a reveal that landed survives
// one; a row that moved does not keep its selection. So a pass follows rows
// drawn again while the last reveal is not known to have landed, and follows
// the caret's row losing its selection once it has. Neither applies once the
// user has moved the tree: a pass then would pull it back to the editor.

import * as vscode from 'vscode';
import { debounce } from './util/debounce';
import type { VbaCaretPosition } from './vbaCaretProcedure';
import type { XlideNode } from './projectExplorer';

/** What following needs from the explorer. */
export interface FollowedExplorer {
    /** Rows drawn again: one module's rows, or (undefined) every row. */
    readonly onDidReplaceRows: vscode.Event<{ filePath: string; moduleName: string } | undefined>;
    resolveModuleNode(filePath: string, moduleName: string): Promise<XlideNode | undefined>;
    resolveProcedureNode(filePath: string, moduleName: string, label: string): Promise<XlideNode | undefined>;
    setActiveModule(filePath: string, moduleName: string): void;
    clearActiveModule(filePath: string, moduleName: string): void;
    collapseAllFolders(): void;
    notifyFolderExpansion(node: XlideNode, expanded: boolean): void;
    notifyProjectCollapsed(filePath: string): void;
    getParent(node: XlideNode): XlideNode | undefined;
    /** A row's identity, the same for every node object that draws it. */
    rowIdentity(node: XlideNode): string;
}

export interface ExplorerFollowDeps {
    explorer: FollowedExplorer;
    treeView: Pick<vscode.TreeView<XlideNode>,
        | 'visible' | 'selection' | 'reveal'
        | 'onDidChangeVisibility' | 'onDidChangeSelection' | 'onDidExpandElement' | 'onDidCollapseElement'>;
    caret: { readonly current: VbaCaretPosition | undefined; readonly onDidChange: vscode.Event<unknown> };
    /** xlide.explorer.autoExpandCollapse, read each time. */
    enabled: () => boolean;
    /** The modules whose last tab a tab change closed. */
    modulesClosedBy?: (event: vscode.TabChangeEvent) => Array<{ projectPath: string; moduleName: string }>;
}

/** Long enough for a burst of tab switches to settle into one pass. */
export const FOLLOW_QUIET_MS = 60;

/**
 * How long after a reveal settles its expansions are still its own. VS Code
 * delivers the expand events separately from the reveal's promise.
 */
export const REVEAL_EVENT_GRACE_MS = 1500;

/**
 * 'pending': a pass is due, or the last one has not been seen to land.
 * 'landed': the tree shows the caret's row selected. 'userLed': the user
 * moved the tree since the editor last moved.
 */
type FollowState = 'pending' | 'landed' | 'userLed';

export class ExplorerFollow implements vscode.Disposable {
    private _generation = 0;
    private _queue: Promise<void> = Promise.resolve();
    /** Rows a reveal of ours expands or selects: identity to when that stops being ours. */
    private readonly _ours = new Map<string, number>();
    private _state: FollowState = 'pending';
    /** The row the latest pass revealed. */
    private _target: string | undefined;
    private readonly _disposables: vscode.Disposable[] = [];
    private readonly _run = debounce(() => {
        const generation = this._generation;
        this._queue = this._queue.then(() => this._pass(generation)).catch(() => undefined);
    }, FOLLOW_QUIET_MS);

    constructor(private readonly _deps: ExplorerFollowDeps) {
        const { treeView, caret, explorer } = _deps;
        this._disposables.push(
            caret.onDidChange(() => this.schedule()),
            treeView.onDidChangeVisibility((event) => {
                if (event.visible) {
                    this.schedule();
                }
            }),
            explorer.onDidReplaceRows(() => {
                if (this._state === 'pending') {
                    this.schedule();
                }
            }),
            treeView.onDidChangeSelection((event) => this._selectionChanged(event.selection)),
            // A row expanded by hand: a module becomes the one the accordion
            // keeps open, and a folder keeps the state it was given until the
            // editor moves to another folder. A row our own reveal expanded
            // is neither.
            treeView.onDidExpandElement((event) => {
                if (!_deps.enabled() || this._isOurs(event.element)) {
                    return;
                }
                if (event.element.kind === 'module' && event.element.moduleName) {
                    explorer.setActiveModule(event.element.filePath, event.element.moduleName);
                    this._state = 'userLed';
                }
                explorer.notifyFolderExpansion(event.element, true);
            }),
            treeView.onDidCollapseElement((event) => {
                // A project folded by hand stays folded: a refresh must not
                // spring it open again because it holds the active module.
                if (event.element.kind === 'project') {
                    explorer.notifyProjectCollapsed(event.element.filePath);
                }
                if (_deps.enabled()) {
                    explorer.notifyFolderExpansion(event.element, false);
                }
            }),
            vscode.window.onDidChangeActiveTextEditor((editor) => {
                // Nothing open at all is the last editor closing, not focus
                // moving to a panel: no module is being edited, so the folder
                // layout goes back to its resting shape.
                if (!editor && vscode.window.visibleTextEditors.length === 0 && _deps.enabled()) {
                    explorer.collapseAllFolders();
                }
            }),
            vscode.window.tabGroups.onDidChangeTabs((event) => {
                if (!_deps.enabled() || !_deps.modulesClosedBy) {
                    return;
                }
                const closed = _deps.modulesClosedBy(event);
                for (const location of closed) {
                    explorer.clearActiveModule(location.projectPath, location.moduleName);
                }
                if (closed.length > 0) {
                    this.schedule();
                }
            }),
            vscode.workspace.onDidChangeConfiguration((event) => {
                if (event.affectsConfiguration('xlide.explorer.autoExpandCollapse') && _deps.enabled()) {
                    this.schedule();
                }
            }),
        );
    }

    /** Asks for a pass; any pass already asked for, or running, is overtaken. */
    schedule(): void {
        this._state = 'pending';
        this._generation += 1;
        this._run();
    }

    private _selectionChanged(selection: readonly XlideNode[]): void {
        if (selection.length === 0) {
            // Nobody chose another row, yet the caret's row is no longer
            // selected: it was drawn again under another parent or name.
            if (this._state === 'landed') {
                this.schedule();
            }
            return;
        }
        const row = this._deps.explorer.rowIdentity(selection[0]);
        if (row === this._target) {
            if (this._state === 'pending') {
                this._state = 'landed';
            }
        } else if (!this._isOurs(selection[0])) {
            this._state = 'userLed';
        }
    }

    /**
     * One pass: the module and procedure the caret is in, revealed. Stops at
     * the next step once a newer request exists, so a late finish never lands
     * on a module the editor has left.
     */
    private async _pass(generation: number): Promise<void> {
        const stale = (): boolean => generation !== this._generation;
        const { explorer, treeView, caret, enabled } = this._deps;
        if (stale() || !enabled()) {
            return;
        }
        const position = caret.current;
        if (!position) {
            // Not a module in front - the Output panel, a settings page. The
            // tree stays as it is: focus moving away folds nothing.
            return;
        }
        this._state = 'pending';
        this._target = undefined;
        // Loaded before the module is made active: which folders to open
        // comes from the module's row.
        const moduleNode = await explorer.resolveModuleNode(position.projectPath, position.moduleName);
        if (stale()) {
            return;
        }
        if (moduleNode) {
            // Making the module active redraws it and its folders open, and
            // VS Code can report that as it reports a click, even when this
            // pass goes stale before its reveal.
            this._claim(moduleNode, Date.now() + REVEAL_EVENT_GRACE_MS);
        }
        explorer.setActiveModule(position.projectPath, position.moduleName);
        if (!moduleNode || !treeView.visible) {
            // A hidden tree catches up when it is shown again, and a module
            // it does not list yet when its rows are next drawn again.
            return;
        }
        const procedureNode = position.procedure
            ? await explorer.resolveProcedureNode(position.projectPath, position.moduleName, position.label)
            : undefined;
        if (stale()) {
            return;
        }
        // The procedure's row, which opens every row above it; or, in the
        // declarations section, the module itself, opened.
        const target = procedureNode ?? moduleNode;
        this._target = explorer.rowIdentity(target);
        await this._reveal(target, target === moduleNode);
        const selection = treeView.selection;
        if (!stale() && this._state === 'pending'
            && selection.length === 1 && explorer.rowIdentity(selection[0]) === this._target) {
            this._state = 'landed';
        }
    }

    private async _reveal(node: XlideNode, expand: boolean): Promise<void> {
        // The row itself and every row above it: the rows the reveal selects
        // and expands. A procedure row cannot expand, so counting it among
        // them only marks its selection as ours.
        const rows = this._claim(node, Number.POSITIVE_INFINITY);
        try {
            await this._deps.treeView.reveal(node, { select: true, focus: false, expand });
        } catch {
            // Whether it landed is read from the selection; VS Code drops
            // most reveals it cannot carry out without rejecting them.
        } finally {
            const until = Date.now() + REVEAL_EVENT_GRACE_MS;
            for (const row of rows) {
                this._ours.set(row, until);
            }
        }
    }

    /** Marks a row and every row above it as ours until then; returns their identities. */
    private _claim(node: XlideNode, until: number): string[] {
        const now = Date.now();
        for (const [row, expires] of this._ours) {
            if (expires < now) {
                this._ours.delete(row);
            }
        }
        const rows: string[] = [];
        for (let row: XlideNode | undefined = node; row; row = this._deps.explorer.getParent(row)) {
            const identity = this._deps.explorer.rowIdentity(row);
            rows.push(identity);
            this._ours.set(identity, Math.max(this._ours.get(identity) ?? 0, until));
        }
        return rows;
    }

    /** Whether an expansion or a selection is one a pass of ours caused. */
    private _isOurs(node: XlideNode): boolean {
        const row = this._deps.explorer.rowIdentity(node);
        const until = this._ours.get(row);
        if (until === undefined) {
            return false;
        }
        if (until < Date.now()) {
            this._ours.delete(row);
            return false;
        }
        return true;
    }

    dispose(): void {
        this._run.dispose();
        for (const disposable of this._disposables) {
            disposable.dispose();
        }
    }
}
