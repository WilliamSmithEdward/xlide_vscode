import { beforeEach, describe, expect, it, vi } from 'vitest';

const state = vi.hoisted(() => ({
    projects: [] as Array<{ fsPath: string }>,
    /** The project of the module in front, or none. */
    activeProject: undefined as string | undefined,
}));

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());

vi.mock('../src/macroContainerDiscovery', () => ({
    findMacroContainerFiles: vi.fn(async () => state.projects),
}));

// The module address encoding has its own tests; here it only has to name a
// project, so the two readers are stubbed rather than a URI hand-built.
vi.mock('../src/xlideFileSystem', () => ({
    XLIDE_SCHEME: 'xlide-vba',
    activeLocalVbaEditor: () => (state.activeProject
        ? { document: { uri: { scheme: 'xlide-vba' } } }
        : undefined),
    decodeModuleUri: () => ({ projectPath: state.activeProject, moduleName: 'Module1' }),
}));

import {
    forgetLastOpenedProject,
    noteProjectOpened,
    resolveProjectTarget,
    SELECTED_PROJECT_STATE_KEY,
} from '../src/projectTarget';

// The ladder that decides which file a command with no argument acts on.
// A keybinding carries no tree row and no editor, so this is the whole of
// what it knows - and opening the wrong workbook in Excel is worse than
// opening none, which is why the bottom of the ladder is nothing rather
// than a guess.

const BOOK = 'C:\\work\\Book.xlsm';
const OTHER = 'C:\\work\\Other.xlsm';

const uri = (fsPath: string) => ({ fsPath });

/** An open XLIDE module document of the given project. */
function editorOn(projectPath: string): void {
    state.activeProject = projectPath;
}

/** Workspace state holding the sidebar's pick, or nothing. */
const stateWith = (selected?: string) => ({
    get: (key: string) => (key === SELECTED_PROJECT_STATE_KEY ? selected : undefined),
}) as never;

describe('which project a command with no argument acts on', () => {
    beforeEach(() => {
        state.projects = [uri(BOOK), uri(OTHER)];
        state.activeProject = undefined;
        forgetLastOpenedProject();
    });

    it('does nothing when several files are open to it and none is indicated', async () => {
        // The whole point: no row, no editor, nothing opened yet. A guess here
        // opens someone else's workbook.
        expect(await resolveProjectTarget()).toBeUndefined();
    });

    it('takes the only file in the workspace', async () => {
        state.projects = [uri(BOOK)];
        expect(await resolveProjectTarget()).toEqual({ filePath: BOOK, source: 'singleProject' });
    });

    it('takes the file the active module belongs to', async () => {
        editorOn(OTHER);
        expect(await resolveProjectTarget()).toEqual({ filePath: OTHER, source: 'activeEditor' });
    });

    it('takes the file last worked in once the editor has moved on', async () => {
        noteProjectOpened(OTHER);
        state.activeProject = undefined;
        expect(await resolveProjectTarget()).toEqual({ filePath: OTHER, source: 'lastOpened' });
    });

    it('lets the active editor outrank what was opened before it', async () => {
        noteProjectOpened(OTHER);
        editorOn(BOOK);
        expect(await resolveProjectTarget()).toEqual({ filePath: BOOK, source: 'activeEditor' });
    });

    it('lets the sidebar pick outrank both: it is the one explicit choice', async () => {
        noteProjectOpened(BOOK);
        editorOn(BOOK);
        expect(await resolveProjectTarget({ workspaceState: stateWith(OTHER) }))
            .toEqual({ filePath: OTHER, source: 'sidebarSelection' });
    });

    it('ignores a remembered file that is no longer in the workspace', async () => {
        // Deleted, renamed, or the folder closed. Answering with a path that
        // is not there would open a dialog complaining about it.
        noteProjectOpened('C:\\work\\Gone.xlsm');
        expect(await resolveProjectTarget()).toBeUndefined();
    });

    it('ignores a sidebar pick that is no longer in the workspace', async () => {
        expect(await resolveProjectTarget({ workspaceState: stateWith('C:\\work\\Gone.xlsm') }))
            .toBeUndefined();
    });

    it('falls to the only file when the remembered one went', async () => {
        state.projects = [uri(BOOK)];
        noteProjectOpened('C:\\work\\Gone.xlsm');
        expect(await resolveProjectTarget()).toEqual({ filePath: BOOK, source: 'singleProject' });
    });

    it('answers nothing when the workspace has no macro-enabled file at all', async () => {
        state.projects = [];
        editorOn(BOOK);
        expect(await resolveProjectTarget()).toBeUndefined();
    });

    it('uses the caller s own project list rather than globbing again', async () => {
        const discovery = await import('../src/macroContainerDiscovery');
        vi.mocked(discovery.findMacroContainerFiles).mockClear();
        expect(await resolveProjectTarget({ projects: [uri(OTHER)] as never }))
            .toEqual({ filePath: OTHER, source: 'singleProject' });
        expect(discovery.findMacroContainerFiles).not.toHaveBeenCalled();
    });
});
