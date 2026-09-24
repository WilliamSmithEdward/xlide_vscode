import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

const host = vi.hoisted(() => ({
    tabsChanged: undefined as undefined | ((event: unknown) => void),
}));

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
    window: {
        onDidChangeActiveTextEditor: vi.fn(() => ({ dispose: () => undefined })),
        visibleTextEditors: [],
        tabGroups: {
            all: [],
            onDidChangeTabs: vi.fn((listener: (event: unknown) => void) => {
                host.tabsChanged = listener;
                return { dispose: () => undefined };
            }),
        },
    },
}));

import { ExplorerFollow, FOLLOW_QUIET_MS, REVEAL_EVENT_GRACE_MS, type ExplorerFollowDeps } from '../src/explorerFollow';
import type { XlideNode } from '../src/projectExplorer';
import type { VbaCaretPosition } from '../src/vbaCaretProcedure';
import { EventEmitter } from './helpers/vscodeMock';

const PROJECT = 'C:\\work\\Book.xlsm';
const project: XlideNode = { kind: 'project', label: 'Book.xlsm', filePath: PROJECT };

/** Project Book.xlsm with modules A and B, each holding Sub <name>First and Sub <name>Second. */
function world(options: { enabled?: () => boolean } = {}) {
    const modules = new Map<string, XlideNode>();
    const procedures = new Map<string, XlideNode>();
    for (const name of ['A', 'B']) {
        modules.set(name, { kind: 'module', label: name, filePath: PROJECT, moduleName: name });
        for (const which of ['First', 'Second']) {
            const label = `Sub ${name}${which}`;
            procedures.set(label, { kind: 'sub', label, filePath: PROJECT, moduleName: name });
        }
    }
    const listed = new Set(['A', 'B']);
    const rowsReplaced = new EventEmitter<{ filePath: string; moduleName: string } | undefined>();
    const explorer = {
        onDidReplaceRows: rowsReplaced.event,
        resolveModuleNode: vi.fn(async (_path: string, name: string) => (listed.has(name) ? modules.get(name) : undefined)),
        resolveProcedureNode: vi.fn(async (_path: string, _name: string, label: string) => procedures.get(label)),
        setActiveModule: vi.fn(),
        clearActiveModule: vi.fn(),
        foldModuleUnlessActive: vi.fn(),
        collapseAllFolders: vi.fn(),
        notifyFolderExpansion: vi.fn(),
        notifyProjectCollapsed: vi.fn(),
        getParent: (node: XlideNode): XlideNode | undefined =>
            node.kind === 'sub' ? modules.get(node.moduleName ?? '') : node.kind === 'module' ? project : undefined,
        rowIdentity: (node: XlideNode): string => `${node.kind}::${node.moduleName ?? ''}::${node.label}`,
    };

    const visibility = new EventEmitter<{ visible: boolean }>();
    const selectionChanged = new EventEmitter<{ selection: readonly XlideNode[] }>();
    const expanded = new EventEmitter<{ element: XlideNode }>();
    const collapsed = new EventEmitter<{ element: XlideNode }>();
    const treeView = {
        visible: true,
        selection: [] as readonly XlideNode[],
        /** Off: reveals VS Code drops without a word, as it does for rows a refresh replaced. */
        revealsLand: true,
        reveal: vi.fn(async (node: XlideNode, _options?: unknown) => {
            if (treeView.revealsLand) {
                treeView.select([node]);
            }
        }),
        select(selection: readonly XlideNode[]): void {
            treeView.selection = selection;
            selectionChanged.fire({ selection });
        },
        show(): void {
            treeView.visible = true;
            visibility.fire({ visible: true });
        },
        onDidChangeVisibility: visibility.event,
        onDidChangeSelection: selectionChanged.event,
        onDidExpandElement: expanded.event,
        onDidCollapseElement: collapsed.event,
    };

    const caretChanged = new EventEmitter<unknown>();
    const caret = {
        current: undefined as VbaCaretPosition | undefined,
        onDidChange: caretChanged.event,
        /** The caret in a procedure of the module, or in its declarations. */
        moveTo(moduleName: string, which?: 'First' | 'Second'): void {
            caret.current = {
                projectPath: PROJECT,
                moduleName,
                native: false,
                procedure: which ? ({ name: `${moduleName}${which}` } as never) : undefined,
                label: which ? `Sub ${moduleName}${which}` : '(Declarations)',
            };
            caretChanged.fire(caret.current);
        },
    };

    const follow = new ExplorerFollow({
        explorer,
        treeView: treeView as unknown as ExplorerFollowDeps['treeView'],
        caret,
        enabled: options.enabled ?? (() => true),
        modulesClosedBy: () => [{ projectPath: PROJECT, moduleName: 'A' }],
    });
    return {
        modules, procedures, listed, explorer, rowsReplaced, treeView, caret, follow,
        expand: (element: XlideNode) => expanded.fire({ element }),
        revealed: () => treeView.reveal.mock.calls.map(([node]) => node.label),
    };
}

/** Past the quiet period, with every await of the pass it started run out. */
async function settle(): Promise<void> {
    await vi.advanceTimersByTimeAsync(FOLLOW_QUIET_MS + 1);
    for (let i = 0; i < 5; i++) {
        await new Promise((resolve) => setImmediate(resolve));
    }
}

describe('the explorer following the editor', () => {
    let follows: ExplorerFollow[] = [];

    beforeEach(() => {
        vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'Date'] });
        follows = [];
    });

    afterEach(() => {
        follows.forEach((follow) => follow.dispose());
        vi.useRealTimers();
    });

    function make(options?: { enabled?: () => boolean }) {
        const made = world(options);
        follows.push(made.follow);
        return made;
    }

    it('reveals the procedure the caret is in, once the switching stops', async () => {
        const { caret, explorer, revealed, treeView } = make();
        for (const [name, which] of [['A', 'First'], ['B', 'First'], ['A', 'Second'], ['B', 'Second']] as const) {
            caret.moveTo(name, which);
            await vi.advanceTimersByTimeAsync(FOLLOW_QUIET_MS / 2);
        }
        await settle();

        expect(explorer.resolveModuleNode).toHaveBeenCalledTimes(1);
        expect(explorer.setActiveModule.mock.calls).toEqual([[PROJECT, 'B']]);
        expect(revealed()).toEqual(['Sub BSecond']);
        expect(treeView.reveal).toHaveBeenLastCalledWith(expect.anything(), { select: true, focus: false, expand: false });
    });

    it('opens the module itself when the caret is above its first procedure', async () => {
        const { caret, modules, treeView } = make();
        caret.moveTo('A');
        await settle();

        expect(treeView.reveal).toHaveBeenCalledWith(modules.get('A'), { select: true, focus: false, expand: true });
    });

    it('stops a pass the editor has moved on from, instead of finishing it late', async () => {
        const { caret, explorer, revealed } = make();
        let release: () => void = () => undefined;
        explorer.resolveModuleNode.mockImplementationOnce(async (_path: string, name: string) => {
            await new Promise<void>((resolve) => { release = resolve; });
            return { kind: 'module', label: name, filePath: PROJECT, moduleName: name };
        });
        caret.moveTo('A', 'First');
        await settle();
        // A is still loading its rows when the editor moves to B.
        caret.moveTo('B', 'First');
        await settle();
        release();
        await settle();

        expect(explorer.setActiveModule.mock.calls).toEqual([[PROJECT, 'B']]);
        expect(revealed()).toEqual(['Sub BFirst']);
    });

    it('does not take the expansions its own reveal causes for clicks', async () => {
        const { caret, explorer, expand, modules } = make();
        caret.moveTo('A', 'First');
        await settle();
        caret.moveTo('B', 'First');
        await settle();

        // VS Code reports the reveal's expansions after it; A's arrives late.
        expand(modules.get('A')!);
        expand(modules.get('B')!);
        expect(explorer.setActiveModule.mock.calls).toEqual([[PROJECT, 'A'], [PROJECT, 'B']]);

        // Well after the reveal, the same event is somebody's click.
        await vi.advanceTimersByTimeAsync(REVEAL_EVENT_GRACE_MS + 1);
        expand(modules.get('A')!);
        expect(explorer.setActiveModule).toHaveBeenLastCalledWith(PROJECT, 'A');
    });

    it('folds a module its reveal opened after the editor had left it', async () => {
        // The module's tab closed while the reveal was under way: the
        // accordion folded the module, and the reveal then opened it again.
        // Taken as the reveal's own, the expansion stayed, with no tab.
        const { caret, explorer, expand, modules, treeView } = make();
        let release: () => void = () => undefined;
        treeView.reveal.mockImplementationOnce(async (node: XlideNode) => {
            await new Promise<void>((resolve) => { release = resolve; });
            treeView.select([node]);
        });
        caret.moveTo('A', 'First');
        await settle();
        caret.current = undefined;
        host.tabsChanged?.({ closed: [], opened: [], changed: [] });
        release();
        await settle();

        expand(modules.get('A')!);

        expect(explorer.clearActiveModule).toHaveBeenCalledWith(PROJECT, 'A');
        expect(explorer.foldModuleUnlessActive).toHaveBeenCalledWith(PROJECT, 'A');
        expect(explorer.setActiveModule.mock.calls).toEqual([[PROJECT, 'A']]);
    });

    it('claims the rows it opens before the reveal, so a pass overtaken in between leaves no clicks behind', async () => {
        const { caret, explorer, expand, modules } = make();
        let release: () => void = () => undefined;
        explorer.resolveProcedureNode.mockImplementationOnce(async () => {
            await new Promise<void>((resolve) => { release = resolve; });
            return undefined;
        });
        caret.moveTo('A', 'First');
        await settle();
        // A is active and drawn open, and the pass is waiting for its procedures.
        expand(modules.get('A')!);
        caret.moveTo('B', 'First');
        release();
        await settle();

        expect(explorer.setActiveModule.mock.calls).toEqual([[PROJECT, 'A'], [PROJECT, 'B']]);
    });

    it('catches up when the tree is shown again', async () => {
        const { caret, explorer, revealed, treeView } = make();
        treeView.visible = false;
        caret.moveTo('B', 'Second');
        await settle();
        expect(explorer.setActiveModule).toHaveBeenCalledWith(PROJECT, 'B');
        expect(revealed()).toEqual([]);

        treeView.show();
        await settle();
        expect(revealed()).toEqual(['Sub BSecond']);
    });

    it('tries again when rows are drawn again while a reveal did not land', async () => {
        const { caret, revealed, rowsReplaced, treeView } = make();
        treeView.revealsLand = false;
        caret.moveTo('A', 'Second');
        await settle();
        expect(revealed()).toEqual(['Sub ASecond']);

        treeView.revealsLand = true;
        rowsReplaced.fire(undefined);
        await settle();
        expect(revealed()).toEqual(['Sub ASecond', 'Sub ASecond']);

        // Landed now: drawing the rows again keeps them, so nothing repeats.
        rowsReplaced.fire(undefined);
        await settle();
        expect(revealed()).toHaveLength(2);
    });

    it('reveals a module the tree did not list yet once its rows are drawn again', async () => {
        const { caret, explorer, listed, revealed, rowsReplaced } = make();
        listed.delete('B');
        caret.moveTo('B', 'First');
        await settle();
        // The accordion already knows; there is just no row to reveal.
        expect(explorer.setActiveModule).toHaveBeenCalledWith(PROJECT, 'B');
        expect(revealed()).toEqual([]);

        listed.add('B');
        rowsReplaced.fire({ filePath: PROJECT, moduleName: 'B' });
        await settle();
        expect(revealed()).toEqual(['Sub BFirst']);
    });

    it('marks the caret\'s procedure again when a redraw takes its selection away', async () => {
        const { caret, revealed, treeView } = make();
        caret.moveTo('A', 'First');
        await settle();

        treeView.select([]);
        await settle();
        expect(revealed()).toEqual(['Sub AFirst', 'Sub AFirst']);
    });

    it('leaves a tree the user has moved in alone until the editor moves', async () => {
        const { caret, procedures, revealed, rowsReplaced, treeView } = make();
        caret.moveTo('A', 'First');
        await settle();
        await vi.advanceTimersByTimeAsync(REVEAL_EVENT_GRACE_MS + 1);

        // Arrowing down the tree to another row, then rows drawn again.
        treeView.select([procedures.get('Sub BFirst')!]);
        rowsReplaced.fire(undefined);
        treeView.select([]);
        await settle();
        expect(revealed()).toEqual(['Sub AFirst']);

        caret.moveTo('A', 'Second');
        await settle();
        expect(revealed()).toEqual(['Sub AFirst', 'Sub ASecond']);
    });

    it('leaves the tree alone when a module opened by hand moves the accordion', async () => {
        const { caret, explorer, expand, modules, revealed, rowsReplaced, treeView } = make();
        caret.moveTo('A', 'First');
        await settle();
        await vi.advanceTimersByTimeAsync(REVEAL_EVENT_GRACE_MS + 1);

        expand(modules.get('B')!);
        expect(explorer.setActiveModule).toHaveBeenLastCalledWith(PROJECT, 'B');
        // B opening folds A, and the caret's row goes with it.
        treeView.select([]);
        rowsReplaced.fire(undefined);
        await settle();
        expect(revealed()).toEqual(['Sub AFirst']);
        expect(explorer.setActiveModule).toHaveBeenLastCalledWith(PROJECT, 'B');
    });

    it('follows the editor that is left when a module\'s last tab closes', async () => {
        const { caret, explorer, revealed } = make();
        caret.moveTo('A', 'First');
        await settle();
        caret.current = { ...caret.current!, moduleName: 'B', label: 'Sub BFirst' };

        host.tabsChanged?.({ closed: [], opened: [], changed: [] });
        await settle();
        expect(explorer.clearActiveModule).toHaveBeenCalledWith(PROJECT, 'A');
        expect(revealed()).toEqual(['Sub AFirst', 'Sub BFirst']);
    });

    it('does nothing while the setting is off', async () => {
        const { caret, explorer, revealed, rowsReplaced } = make({ enabled: () => false });
        caret.moveTo('A', 'First');
        rowsReplaced.fire(undefined);
        await settle();

        expect(explorer.resolveModuleNode).not.toHaveBeenCalled();
        expect(explorer.setActiveModule).not.toHaveBeenCalled();
        expect(revealed()).toEqual([]);
    });

    it('leaves the tree as it is when no module is in front', async () => {
        const { caret, explorer, revealed } = make();
        caret.current = undefined;
        host.tabsChanged?.({ closed: [], opened: [], changed: [] });
        await settle();

        expect(explorer.setActiveModule).not.toHaveBeenCalled();
        expect(revealed()).toEqual([]);
    });
});
