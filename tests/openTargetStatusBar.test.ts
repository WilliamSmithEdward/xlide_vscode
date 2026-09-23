import { beforeEach, describe, expect, it, vi } from 'vitest';

const state = vi.hoisted(() => ({
    projects: [] as Array<{ fsPath: string }>,
    item: undefined as undefined | { text: string; tooltip: string; command: string; visible: boolean; name: string },
    editorListeners: [] as Array<() => void>,
}));

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
    StatusBarAlignment: { Left: 1, Right: 2 },
    window: {
        createStatusBarItem: vi.fn(() => {
            const item = {
                text: '', tooltip: '', command: '', name: '', visible: false,
                show() { item.visible = true; },
                hide() { item.visible = false; },
                dispose: vi.fn(),
            };
            state.item = item;
            return item;
        }),
        onDidChangeActiveTextEditor: vi.fn((listener: () => void) => {
            state.editorListeners.push(listener);
            return { dispose: vi.fn() };
        }),
    },
    workspace: {
        onDidChangeWorkspaceFolders: vi.fn(() => ({ dispose: vi.fn() })),
        textDocuments: [],
    },
}));

vi.mock('../src/macroContainerDiscovery', () => ({
    findMacroContainerFiles: vi.fn(async () => state.projects),
}));

vi.mock('../src/xlideFileSystem', () => ({
    XLIDE_SCHEME: 'xlide-vba',
    activeLocalVbaEditor: () => undefined,
    decodeModuleUri: () => ({ projectPath: undefined, moduleName: 'Module1' }),
}));

import { OpenTargetStatusBar, openTargetHint } from '../src/openTargetStatusBar';
import {
    forgetLastOpenedProject,
    noteProjectOpened,
    onDidChangeProjectTargetInputs,
    SELECTED_PROJECT_STATE_KEY,
    storeSelectedProject,
} from '../src/projectTarget';

const BOOK = 'C:\\work\\Book.xlsm';
const DECK = 'C:\\work\\Deck.pptm';

/** Workspace state that remembers what it is given. */
function memento(): { get: (key: string) => unknown; update: (key: string, value: unknown) => Promise<void> } {
    const values = new Map<string, unknown>();
    return {
        get: (key) => values.get(key),
        update: async (key, value) => { values.set(key, value); },
    };
}

const settle = async (): Promise<void> => {
    for (let i = 0; i < 10; i++) {
        await new Promise((resolve) => setImmediate(resolve));
    }
};

describe('what the hint says', () => {
    it('names the file and why it was chosen', () => {
        expect(openTargetHint(2, { filePath: BOOK, source: 'sidebarSelection' }, 'win32')).toEqual({
            text: '$(link-external) Book.xlsm',
            tooltip: 'Ctrl+Alt+O opens Book.xlsm in Excel (Ctrl+Alt+Shift+O opens it read-only), '
                + 'because it is the file selected in the XLIDE sidebar.',
            command: 'xlide.openInOfficeApp',
        });
        expect(openTargetHint(1, { filePath: DECK, source: 'singleProject' }, 'win32')?.tooltip)
            .toBe('Ctrl+Alt+O opens Deck.pptm in PowerPoint (Ctrl+Alt+Shift+O opens it read-only), because it is the only one in this window.');
        expect(openTargetHint(3, { filePath: BOOK, source: 'lastOpened' }, 'win32')?.tooltip)
            .toContain('because it is the file you last had a module open from.');
        expect(openTargetHint(3, { filePath: BOOK, source: 'activeEditor' }, 'win32')?.tooltip)
            .toContain('because a module of it is the editor in front.');
    });

    it('says ambiguous when the key would open nothing, and where to fix that', () => {
        expect(openTargetHint(3, undefined, 'win32')).toEqual({
            text: '$(link-external) ambiguous',
            tooltip: 'Ctrl+Alt+O opens nothing: this window has 3 files, and nothing says which one. '
                + 'Select one in the XLIDE sidebar, or open one of its modules.',
            command: 'xlide.sidebar.focus',
        });
    });

    it('names the keys as the Mac keybinding has them', () => {
        expect(openTargetHint(1, { filePath: BOOK, source: 'singleProject' }, 'darwin')?.tooltip)
            .toMatch(/^Cmd\+Alt\+O opens Book\.xlsm in Excel \(Cmd\+Alt\+Shift\+O opens it read-only\)/);
    });

    it('hides when the window has nothing to open', () => {
        expect(openTargetHint(0, undefined, 'win32')).toBeUndefined();
    });
});

describe('the ladder says when it may have moved', () => {
    beforeEach(() => {
        forgetLastOpenedProject();
    });

    it('fires for a new project worked in, a pick in the sidebar and a forget, and not for the same project again', async () => {
        let fired = 0;
        const subscription = onDidChangeProjectTargetInputs(() => { fired += 1; });
        noteProjectOpened(BOOK);
        noteProjectOpened(BOOK);
        expect(fired).toBe(1);
        const workspaceState = memento();
        await storeSelectedProject(workspaceState as never, DECK);
        expect(workspaceState.get(SELECTED_PROJECT_STATE_KEY)).toBe(DECK);
        expect(fired).toBe(2);
        forgetLastOpenedProject();
        expect(fired).toBe(3);
        subscription.dispose();
    });
});

describe('the status bar item', () => {
    beforeEach(() => {
        forgetLastOpenedProject();
        state.editorListeners = [];
    });

    it('follows the pick in the sidebar, and says ambiguous once nothing points anywhere', async () => {
        state.projects = [{ fsPath: BOOK }, { fsPath: DECK }];
        const workspaceState = memento();
        const bar = new OpenTargetStatusBar(workspaceState as never);
        await settle();
        expect(state.item).toMatchObject({ text: '$(link-external) ambiguous', visible: true, command: 'xlide.sidebar.focus' });
        expect(state.item?.name).toBe('XLIDE: Where Open in Office Application Goes');

        await storeSelectedProject(workspaceState as never, DECK);
        await settle();
        expect(state.item).toMatchObject({ text: '$(link-external) Deck.pptm', command: 'xlide.openInOfficeApp' });

        await storeSelectedProject(workspaceState as never, undefined);
        noteProjectOpened(BOOK);
        await settle();
        expect(state.item?.text).toBe('$(link-external) Book.xlsm');
        expect(state.item?.tooltip).toContain('the file you last had a module open from');
        bar.dispose();
    });

    it('hides in a window with no file to open', async () => {
        state.projects = [];
        const bar = new OpenTargetStatusBar(memento() as never);
        await settle();
        expect(state.item?.visible).toBe(false);
        bar.dispose();
    });
});
