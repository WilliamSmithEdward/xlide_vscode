// The explorer following the editor, issue #66: the folder layout (the same
// modules the flat tree lists, placed by their `@Folder` annotation), the
// folders opening on the way to the module being edited, and the row for the
// procedure the caret is in.

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

const vscodeMock = vi.hoisted(() => ({
    findFiles: vi.fn(),
    showErrorMessage: vi.fn(),
    treeEvents: [] as unknown[],
}));

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
    EventEmitter: class {
        event = vi.fn();
        fire = vi.fn((node?: unknown) => {
            vscodeMock.treeEvents.push(node);
        });
    },
    window: {
        showErrorMessage: vscodeMock.showErrorMessage,
    },
    workspace: {
        findFiles: vscodeMock.findFiles,
        workspaceFolders: [{ uri: { fsPath: 'C:\\work' } }],
    },
}));

import { ProjectExplorer, type XlideNode } from '../src/projectExplorer';
import { AGENT_REVIEW_COLOR_ID, parseDecorationUri } from '../src/agentReviewDecorations';
import { keepAgentChange, presentAgentModuleWrite } from '../src/xlideAgentDiff';

const BOOK = 'C:\\work\\Book.xlsm';

/**
 * A workbook shaped like the vbide's FolderFixture: nested folders, a folder
 * name that repeats under a different parent, and modules at the root.
 */
const MODULES = [
    { name: 'ThisWorkbook', type: 'document' },
    { name: 'Sheet1', type: 'document', folder: 'Accounts' },
    { name: 'ReminderForm', type: 'userform', folder: 'Accounts.Billing.Reminders' },
    { name: 'Bare', type: 'standard', folder: 'Accounts.Billing.Reminders' },
    { name: 'Invoice', type: 'class', folder: 'Accounts.Billing' },
    { name: 'Ledger', type: 'standard', folder: 'Accounts.Ledger' },
    { name: 'Posting', type: 'class', folder: 'Accounts.Ledger' },
    { name: 'Helpers', type: 'standard', folder: 'Shared' },
    { name: 'Loose', type: 'standard' },
];

function fakeBridge(modules = MODULES) {
    return {
        call: vi.fn((method: string) => {
            // Fresh entries per call, as the real engine hands back: a test
            // that moves a module must not edit the next test's fixture.
            if (method === 'listModules') { return Promise.resolve(modules.map((m) => ({ ...m }))); }
            if (method === 'listSubs') { return Promise.resolve([]); }
            return Promise.resolve({ isPasswordProtected: false, isSigned: false });
        }),
    } as unknown as ConstructorParameters<typeof ProjectExplorer>[0];
}

async function foldersExplorer(modules = MODULES): Promise<ProjectExplorer> {
    const explorer = new ProjectExplorer(fakeBridge(modules));
    explorer.setView('folders');
    return explorer;
}

/** Walk the whole project as "Folder/Module" rows, in the order drawn. */
async function rows(explorer: ProjectExplorer, node: XlideNode, prefix = ''): Promise<string[]> {
    const out: string[] = [];
    for (const child of await explorer.getChildren(node)) {
        if (child.kind === 'folder') {
            out.push(`${prefix}${child.label}/`);
            out.push(...await rows(explorer, child, `${prefix}${child.label}/`));
        } else if (child.kind === 'module') {
            out.push(`${prefix}${child.label}`);
        }
    }
    return out;
}

beforeEach(() => {
    vscodeMock.findFiles.mockReset();
    vscodeMock.showErrorMessage.mockReset();
    vscodeMock.treeEvents = [];
    vscodeMock.findFiles.mockResolvedValue([{ scheme: 'file', fsPath: BOOK }]);
});

describe('drawing the layout', () => {
    it('nests the modules by annotation, folders first and modules after', async () => {
        const explorer = await foldersExplorer();
        const [project] = await explorer.getChildren();
        expect(await rows(explorer, project)).toEqual([
            'Accounts/',
            'Accounts/Billing/',
            'Accounts/Billing/Reminders/',
            'Accounts/Billing/Reminders/ReminderForm',
            'Accounts/Billing/Reminders/Bare',
            'Accounts/Billing/Invoice',
            'Accounts/Ledger/',
            'Accounts/Ledger/Ledger',
            'Accounts/Ledger/Posting',
            'Accounts/Sheet1',
            'Shared/',
            'Shared/Helpers',
            'ThisWorkbook',
            'Loose',
        ]);
    });

    it('keeps the flat tree order inside a folder and at the root', async () => {
        // Documents, then forms, then standard, then class, each group by name
        // - the flat tree's own order, which the folder layout does not re-sort.
        const flat = new ProjectExplorer(fakeBridge());
        const [flatProject] = await flat.getChildren();
        expect((await flat.getChildren(flatProject)).map((n) => n.label))
            .toEqual(['Sheet1', 'ThisWorkbook', 'ReminderForm', 'Bare', 'Helpers', 'Ledger', 'Loose', 'Invoice', 'Posting']);

        const explorer = await foldersExplorer();
        const [project] = await explorer.getChildren();
        // ReminderForm before Bare is that order, not the alphabetical one.
        expect(await rows(explorer, project)).toContain('Accounts/Billing/Reminders/ReminderForm');
        expect((await explorer.getChildren(project)).filter((n) => n.kind === 'module').map((n) => n.label))
            .toEqual(['ThisWorkbook', 'Loose']);
    });

    it('says how many modules a folder holds, however deep', async () => {
        const explorer = await foldersExplorer();
        const [project] = await explorer.getChildren();
        const [accounts] = await explorer.getChildren(project);
        expect(explorer.getTreeItem(accounts).description).toBe('6 modules');
        const [billing] = await explorer.getChildren(accounts);
        expect(explorer.getTreeItem(billing).description).toBe('3 modules');
        const [, shared] = (await explorer.getChildren(project)).filter((n) => n.kind === 'folder');
        expect(explorer.getTreeItem(shared).description).toBe('1 module');
    });

    it('tooltips a folder with the annotation that makes it', async () => {
        const explorer = await foldersExplorer();
        const [project] = await explorer.getChildren();
        const [accounts] = await explorer.getChildren(project);
        const [billing] = await explorer.getChildren(accounts);
        expect(explorer.getTreeItem(billing).tooltip).toBe('@Folder("Accounts.Billing")');
    });

    it('draws the flat list again when the view goes back to the tree', async () => {
        const explorer = await foldersExplorer();
        const [project] = await explorer.getChildren();
        explorer.setView('tree');
        const [freshProject] = await explorer.getChildren();
        expect((await explorer.getChildren(freshProject)).every((n) => n.kind === 'module')).toBe(true);
        expect(project.filePath).toBe(BOOK);
    });

    it('shows the retry placeholder rather than an empty folder layout', async () => {
        const bridge = {
            call: vi.fn((method: string) => method === 'listModules'
                ? Promise.reject(new Error('sharing violation'))
                : Promise.resolve({ isPasswordProtected: false, isSigned: false })),
        } as unknown as ConstructorParameters<typeof ProjectExplorer>[0];
        const explorer = new ProjectExplorer(bridge);
        explorer.setView('folders');
        const [project] = await explorer.getChildren();
        expect(await explorer.getChildren(project)).toMatchObject([{ kind: 'loadError' }]);
    });
});

describe('finding a node again', () => {
    it('walks a module back up through its folders, which reveal() needs', async () => {
        const explorer = await foldersExplorer();
        const [project] = await explorer.getChildren();
        const [accounts] = await explorer.getChildren(project);
        const [billing] = await explorer.getChildren(accounts);
        const [reminders] = await explorer.getChildren(billing);
        const [form] = await explorer.getChildren(reminders);

        expect(explorer.getParent(form)).toBe(reminders);
        expect(explorer.getParent(reminders)).toBe(billing);
        expect(explorer.getParent(billing)).toBe(accounts);
        expect(explorer.getParent(accounts)).toBe(project);
    });

    it('walks an unannotated module straight up to its project', async () => {
        const explorer = await foldersExplorer();
        const [project] = await explorer.getChildren();
        const loose = (await explorer.getChildren(project)).find((n) => n.label === 'Loose')!;
        expect(explorer.getParent(loose)).toBe(project);
    });

    it('knows a nested folder before anything expanded it', async () => {
        // reveal() walks a module up through folders that were never drawn -
        // opening a module while the tree is collapsed is the ordinary case -
        // so expanding the project alone has to name every folder in it.
        const explorer = await foldersExplorer();
        const [project] = await explorer.getChildren();
        await explorer.getChildren(project);

        const deep = explorer.getFolderNode(BOOK, 'Accounts.Billing.Reminders')!;
        expect(deep).toBeDefined();
        expect(explorer.getParent(deep)).toBe(explorer.getFolderNode(BOOK, 'Accounts.Billing'));
    });
});

describe('an annotation edited in the editor', () => {
    /** Draw the whole project, then read it back after the module moves. */
    async function moved(moduleName: string, folder: string | undefined): Promise<string[]> {
        const explorer = await foldersExplorer();
        const [project] = await explorer.getChildren();
        await rows(explorer, project);
        explorer.setModuleFolder(BOOK, moduleName, folder);
        return rows(explorer, project);
    }

    it('moves the module to the folder the annotation now names', async () => {
        const after = await moved('Helpers', 'Accounts.Ledger');
        expect(after).toContain('Accounts/Ledger/Helpers');
        expect(after).not.toContain('Shared/Helpers');
    });

    it('drops a folder the last module left', async () => {
        // Helpers was the only module in Shared, so Shared has nothing to hold.
        expect(await moved('Helpers', 'Accounts')).not.toContain('Shared/');
    });

    it('makes a folder the annotation names for the first time', async () => {
        expect(await moved('Loose', 'Reports.Monthly')).toEqual(
            expect.arrayContaining(['Reports/', 'Reports/Monthly/', 'Reports/Monthly/Loose']),
        );
    });

    it('sends a module back to the root when the annotation goes away', async () => {
        const after = await moved('Helpers', undefined);
        expect(after).toContain('Helpers');
        expect(after).not.toContain('Shared/');
    });

    it('re-counts the folders the module left and joined', async () => {
        const explorer = await foldersExplorer();
        const [project] = await explorer.getChildren();
        await rows(explorer, project);
        explorer.setModuleFolder(BOOK, 'Helpers', 'Accounts.Ledger');
        await rows(explorer, project);

        expect(explorer.getTreeItem(explorer.getFolderNode(BOOK, 'Accounts.Ledger')!).description)
            .toBe('3 modules');
        expect(explorer.getTreeItem(explorer.getFolderNode(BOOK, 'Accounts')!).description)
            .toBe('7 modules');
    });

    it('redraws nothing when the annotation did not actually change', async () => {
        const explorer = await foldersExplorer();
        const [project] = await explorer.getChildren();
        await rows(explorer, project);
        vscodeMock.treeEvents = [];
        explorer.setModuleFolder(BOOK, 'Helpers', 'Shared');
        expect(vscodeMock.treeEvents).toEqual([]);
    });

    it('survives a refresh, which re-reads the container the edit is not in yet', async () => {
        const explorer = await foldersExplorer();
        const [project] = await explorer.getChildren();
        await rows(explorer, project);
        explorer.setModuleFolder(BOOK, 'Helpers', 'Accounts.Ledger');

        explorer.refresh();
        const [fresh] = await explorer.getChildren();
        expect(await rows(explorer, fresh)).toContain('Accounts/Ledger/Helpers');
    });

    it('goes back to the container once the editor closes', async () => {
        // The saved file is what the listing answers; closing the editor is
        // what makes the tree ask it again.
        const saved = MODULES.map((m) => (m.name === 'Helpers' ? { ...m, folder: 'Shared.Text' } : m));
        let listing = MODULES;
        const bridge = {
            call: vi.fn((method: string) => {
                if (method === 'listModules') { return Promise.resolve(listing.map((m) => ({ ...m }))); }
                if (method === 'listSubs') { return Promise.resolve([]); }
                return Promise.resolve({ isPasswordProtected: false, isSigned: false });
            }),
        } as unknown as ConstructorParameters<typeof ProjectExplorer>[0];

        const explorer = new ProjectExplorer(bridge);
        explorer.setView('folders');
        const [project] = await explorer.getChildren();
        await rows(explorer, project);

        explorer.setModuleFolder(BOOK, 'Helpers', 'Accounts.Ledger');
        expect(await rows(explorer, project)).toContain('Accounts/Ledger/Helpers');

        listing = saved;
        explorer.forgetModuleFolder(BOOK, 'Helpers');
        const [fresh] = await explorer.getChildren();
        expect(await rows(explorer, fresh)).toContain('Shared/Text/Helpers');
    });

    it('asks the container nothing extra for a module nobody edited', async () => {
        const bridge = fakeBridge();
        const explorer = new ProjectExplorer(bridge);
        explorer.setView('folders');
        const [project] = await explorer.getChildren();
        await rows(explorer, project);
        const listings = () => (bridge.call as ReturnType<typeof vi.fn>).mock.calls
            .filter(([method]) => method === 'listModules').length;

        const before = listings();
        explorer.forgetModuleFolder(BOOK, 'Helpers');
        await rows(explorer, project);
        expect(listings()).toBe(before);
    });

    it('still places the module correctly if the flat tree was showing', async () => {
        // The edit lands while the flat list is up; switching to the folder
        // layout has to show where the module went, not where it was.
        const explorer = new ProjectExplorer(fakeBridge());
        const [project] = await explorer.getChildren();
        await explorer.getChildren(project);
        explorer.setModuleFolder(BOOK, 'Helpers', 'Accounts.Ledger');

        explorer.setView('folders');
        const [fresh] = await explorer.getChildren();
        expect(await rows(explorer, fresh)).toContain('Accounts/Ledger/Helpers');
    });
});

describe('the row for the procedure the caret is in', () => {
    function bridgeWithSubs(subs: Array<{ name: string; kind: string; line: number }>) {
        return {
            call: vi.fn((method: string) => {
                if (method === 'listModules') { return Promise.resolve(MODULES.map((m) => ({ ...m }))); }
                if (method === 'listSubs') { return Promise.resolve(subs); }
                return Promise.resolve({ isPasswordProtected: false, isSigned: false });
            }),
        } as unknown as ConstructorParameters<typeof ProjectExplorer>[0];
    }

    async function listed(subs: Array<{ name: string; kind: string; line: number }>) {
        const explorer = new ProjectExplorer(bridgeWithSubs(subs));
        const [project] = await explorer.getChildren();
        const modules = await explorer.getChildren(project);
        const module = modules.find((m) => m.moduleName === 'Helpers')!;
        const rows = await explorer.getChildren(module);
        return { explorer, module, rows };
    }

    it('finds the row by the label the tree draws, whatever the caret spells', async () => {
        const { explorer, rows } = await listed([
            { name: 'Post', kind: 'Sub', line: 1 },
            { name: 'Total', kind: 'Function', line: 9 },
        ]);
        expect(explorer.getProcedureNode(BOOK, 'Helpers', 'Sub Post')).toBe(rows[0]);
        expect(explorer.getProcedureNode(BOOK, 'Helpers', 'function total')).toBe(rows[1]);
    });

    it('keeps Property Get, Let and Set apart, which share a name', async () => {
        const { explorer } = await listed([
            { name: 'Name', kind: 'Property Get', line: 1 },
            { name: 'Name', kind: 'Property Let', line: 5 },
        ]);
        expect(explorer.getProcedureNode(BOOK, 'Helpers', 'Property Get Name')?.line).toBe(1);
        expect(explorer.getProcedureNode(BOOK, 'Helpers', 'Property Let Name')?.line).toBe(5);
    });

    it('hands back the same node every render, which reveal() needs', async () => {
        const { explorer, module, rows } = await listed([{ name: 'Post', kind: 'Sub', line: 1 }]);
        expect(await explorer.getChildren(module)).toEqual(rows);
        expect((await explorer.getChildren(module))[0]).toBe(rows[0]);
    });

    it('names nothing for a procedure the container has not got', async () => {
        // An unsaved rename: the editor says one name, the listing another.
        const { explorer } = await listed([{ name: 'Post', kind: 'Sub', line: 1 }]);
        expect(explorer.getProcedureNode(BOOK, 'Helpers', 'Sub Posted')).toBeUndefined();
        expect(explorer.getProcedureNode(BOOK, 'Nowhere', 'Sub Post')).toBeUndefined();
    });

    it('does not answer with a form designer row', async () => {
        const explorer = new ProjectExplorer({
            call: vi.fn((method: string) => {
                if (method === 'listModules') { return Promise.resolve(MODULES.map((m) => ({ ...m }))); }
                if (method === 'listSubs') { return Promise.resolve([]); }
                return Promise.resolve({ isPasswordProtected: false, isSigned: false });
            }),
        } as unknown as ConstructorParameters<typeof ProjectExplorer>[0]);
        const [project] = await explorer.getChildren();
        const form = (await explorer.getChildren(project)).find((m) => m.moduleName === 'ReminderForm')!;
        const rows = await explorer.getChildren(form);
        expect(rows[0]).toMatchObject({ kind: 'designer', label: 'Designer' });
        expect(explorer.getProcedureNode(BOOK, 'ReminderForm', 'Designer')).toBeUndefined();
    });

    it('does not keep rows a refresh raced, which would outlive every re-fetch', async () => {
        // The listing lands AFTER the refresh that cleared the caches. Keeping
        // the rows built from it would win over the re-fetch for good, since a
        // later render finds them and never looks at the fresh list.
        let release: (subs: Array<{ name: string; kind: string; line: number }>) => void;
        let pending = new Promise<Array<{ name: string; kind: string; line: number }>>((r) => { release = r; });
        const bridge = {
            call: vi.fn((method: string) => {
                if (method === 'listModules') { return Promise.resolve(MODULES.map((m) => ({ ...m }))); }
                if (method === 'listSubs') { return pending; }
                return Promise.resolve({ isPasswordProtected: false, isSigned: false });
            }),
        } as unknown as ConstructorParameters<typeof ProjectExplorer>[0];

        const explorer = new ProjectExplorer(bridge);
        const [project] = await explorer.getChildren();
        const module = (await explorer.getChildren(project)).find((m) => m.moduleName === 'Helpers')!;

        const inFlight = explorer.getChildren(module);
        explorer.refresh();
        release!([{ name: 'Stale', kind: 'Sub', line: 1 }]);
        expect((await inFlight).map((n) => n.label)).toEqual(['Sub Stale']);

        pending = Promise.resolve([{ name: 'Fresh', kind: 'Sub', line: 1 }]);
        const [freshProject] = await explorer.getChildren();
        const freshModule = (await explorer.getChildren(freshProject)).find((m) => m.moduleName === 'Helpers')!;
        expect((await explorer.getChildren(freshModule)).map((n) => n.label)).toEqual(['Sub Fresh']);
    });

    it('forgets the rows when the module is re-listed, so a rename is not stale', async () => {
        const { explorer, module } = await listed([{ name: 'Post', kind: 'Sub', line: 1 }]);
        explorer.refreshModuleSubs(BOOK, 'Helpers');
        expect(explorer.getProcedureNode(BOOK, 'Helpers', 'Sub Post')).toBeUndefined();
        await explorer.getChildren(module);
        expect(explorer.getProcedureNode(BOOK, 'Helpers', 'Sub Post')).toBeDefined();
    });
});

describe('following the editor', () => {
    /** Draw the whole project so every folder node exists, as the tree would. */
    async function drawn(): Promise<{ explorer: ProjectExplorer; project: XlideNode }> {
        const explorer = await foldersExplorer();
        const [project] = await explorer.getChildren();
        await rows(explorer, project);
        return { explorer, project };
    }

    const expanded = (explorer: ProjectExplorer, folder: string): boolean =>
        // 2 is TreeItemCollapsibleState.Expanded.
        explorer.getTreeItem(explorer.getFolderNode(BOOK, folder)!).collapsibleState === 2;

    it('opens every folder on the way to the module being edited', async () => {
        const { explorer } = await drawn();
        explorer.setActiveModule(BOOK, 'Bare');
        expect(expanded(explorer, 'Accounts')).toBe(true);
        expect(expanded(explorer, 'Accounts.Billing')).toBe(true);
        expect(expanded(explorer, 'Accounts.Billing.Reminders')).toBe(true);
        expect(expanded(explorer, 'Accounts.Ledger')).toBe(false);
        expect(expanded(explorer, 'Shared')).toBe(false);
    });

    it('folds the folders it left when the editor moves to another one', async () => {
        const { explorer } = await drawn();
        explorer.setActiveModule(BOOK, 'Bare');
        explorer.setActiveModule(BOOK, 'Helpers');
        expect(expanded(explorer, 'Shared')).toBe(true);
        expect(expanded(explorer, 'Accounts')).toBe(false);
        expect(expanded(explorer, 'Accounts.Billing.Reminders')).toBe(false);
    });

    it('leaves the tree alone while the editor stays inside one folder', async () => {
        const { explorer } = await drawn();
        explorer.setActiveModule(BOOK, 'Ledger');
        explorer.notifyFolderExpansion(explorer.getFolderNode(BOOK, 'Shared')!, true);
        // Posting sits beside Ledger, so the attention has not really moved and
        // the folder opened by hand stays open.
        explorer.setActiveModule(BOOK, 'Posting');
        expect(expanded(explorer, 'Shared')).toBe(true);
        expect(expanded(explorer, 'Accounts.Ledger')).toBe(true);
    });

    it('drops a hand-opened folder once the editor really does move', async () => {
        const { explorer } = await drawn();
        explorer.setActiveModule(BOOK, 'Ledger');
        explorer.notifyFolderExpansion(explorer.getFolderNode(BOOK, 'Shared')!, true);
        explorer.setActiveModule(BOOK, 'Invoice');
        expect(expanded(explorer, 'Shared')).toBe(false);
        expect(expanded(explorer, 'Accounts.Billing')).toBe(true);
    });

    it('keeps a folder the user shut by hand shut, even on the way to the module', async () => {
        const { explorer } = await drawn();
        explorer.setActiveModule(BOOK, 'Bare');
        explorer.notifyFolderExpansion(explorer.getFolderNode(BOOK, 'Accounts.Billing')!, false);
        expect(expanded(explorer, 'Accounts.Billing')).toBe(false);
        expect(expanded(explorer, 'Accounts')).toBe(true);
    });

    it('folds them all when the last editor closes', async () => {
        const { explorer } = await drawn();
        explorer.setActiveModule(BOOK, 'Bare');
        explorer.clearActiveModule(BOOK, 'Bare');
        expect(expanded(explorer, 'Accounts')).toBe(false);
        expect(expanded(explorer, 'Accounts.Billing.Reminders')).toBe(false);
    });

    it('takes the open folders along when the module being edited moves', async () => {
        const { explorer, project } = await drawn();
        explorer.setActiveModule(BOOK, 'Helpers');
        explorer.setModuleFolder(BOOK, 'Helpers', 'Accounts.Ledger');
        await rows(explorer, project);
        expect(expanded(explorer, 'Accounts')).toBe(true);
        expect(expanded(explorer, 'Accounts.Ledger')).toBe(true);
        expect(vscodeMock.treeEvents).toContainEqual({ filePath: BOOK, moduleName: 'Helpers' });

        // And they are the ones folded when the editor moves on.
        explorer.setActiveModule(BOOK, 'Loose');
        expect(expanded(explorer, 'Accounts')).toBe(false);
        expect(expanded(explorer, 'Accounts.Ledger')).toBe(false);
    });

    it('opens nothing when a module nobody is editing moves', async () => {
        const { explorer, project } = await drawn();
        explorer.setActiveModule(BOOK, 'Loose');
        explorer.setModuleFolder(BOOK, 'Helpers', 'Accounts.Ledger');
        await rows(explorer, project);
        expect(expanded(explorer, 'Accounts.Ledger')).toBe(false);
    });

    it('leaves the folders alone for a module the tree has not drawn yet', async () => {
        const explorer = await foldersExplorer();
        const [project] = await explorer.getChildren();
        await rows(explorer, project);
        explorer.setActiveModule(BOOK, 'Bare');
        // Nothing is known about a module that was never listed, and folding
        // the tree the editor is inside would be worse than doing nothing.
        explorer.setActiveModule(BOOK, 'NotListed');
        expect(expanded(explorer, 'Accounts.Billing.Reminders')).toBe(true);
    });
});

describe('an agent edit awaiting review stands out', () => {
    async function agentEdits(moduleName: string): Promise<void> {
        await presentAgentModuleWrite(BOOK, moduleName, {
            before: 'Sub Old()\r\nEnd Sub\r\n',
            beforeExisted: true,
            after: 'Sub New()\r\nEnd Sub\r\n',
        });
    }

    /** A drawn folder layout, every module listed. */
    async function drawnTree() {
        const explorer = await foldersExplorer();
        const [project] = await explorer.getChildren();
        await rows(explorer, project);
        const folder = (path: string) => explorer.getFolderNode(BOOK, path)!;
        return { explorer, project, folder };
    }

    const colourOf = (item: { iconPath?: unknown }) =>
        (item.iconPath as { color?: { id: string } }).color?.id;

    afterEach(() => {
        keepAgentChange(BOOK, 'Ledger');
        keepAgentChange(BOOK, 'Loose');
    });

    it('decorates the project row, so a collapsed tree still shows it', async () => {
        const { explorer, project } = await drawnTree();

        expect(parseDecorationUri(explorer.getTreeItem(project).resourceUri!)).toEqual({
            kind: 'project',
            filePath: BOOK,
        });
        // The path stays the hover text; a decoration URI must not replace it.
        expect(explorer.getTreeItem(project).tooltip).toBe(BOOK);
    });

    it('decorates and tints a pending module, and only while it is pending', async () => {
        const { explorer } = await drawnTree();
        const ledger = explorer.getModuleNode(BOOK, 'Ledger')!;
        expect(explorer.getTreeItem(ledger).resourceUri).toBeUndefined();

        await agentEdits('Ledger');
        const item = explorer.getTreeItem(ledger);

        expect(parseDecorationUri(item.resourceUri!)).toEqual({
            kind: 'module',
            filePath: BOOK,
            moduleName: 'Ledger',
        });
        expect(colourOf(item)).toBe(AGENT_REVIEW_COLOR_ID);
        expect(item.description).toBe('standard ● agent edit');
        // A row with a resourceUri and no tooltip would show the URI on hover.
        expect(String(item.tooltip)).toContain('An agent edited this module');

        keepAgentChange(BOOK, 'Ledger');
        expect(explorer.getTreeItem(ledger).resourceUri).toBeUndefined();
    });

    it('marks every folder a pending module sits under, and no other', async () => {
        const { explorer, folder } = await drawnTree();
        await agentEdits('Ledger');

        for (const path of ['Accounts', 'Accounts.Ledger']) {
            const item = explorer.getTreeItem(folder(path));
            expect(item.description, path).toMatch(/● agent edit$/);
            expect(colourOf(item), path).toBe(AGENT_REVIEW_COLOR_ID);
        }
        for (const path of ['Accounts.Billing', 'Shared']) {
            const item = explorer.getTreeItem(folder(path));
            expect(item.description, path).not.toContain('agent edit');
            expect(colourOf(item), path).toBeUndefined();
        }
        // A folder row never takes a decoration URI: its icon would turn into
        // the file-icon theme's folder.
        expect(explorer.getTreeItem(folder('Accounts')).resourceUri).toBeUndefined();

        keepAgentChange(BOOK, 'Ledger');
        expect(explorer.getTreeItem(folder('Accounts')).description).toBe('6 modules');
    });

    it('marks no folder for a pending module at the project root', async () => {
        const { explorer, folder } = await drawnTree();
        await agentEdits('Loose');

        expect(explorer.getTreeItem(folder('Accounts')).description).toBe('6 modules');
        expect(explorer.getTreeItem(folder('Shared')).description).toBe('1 module');
    });

    it('redraws the module and its folders without reopening or shutting them', async () => {
        const { explorer, folder } = await drawnTree();
        const before = explorer.getTreeItem(folder('Accounts')).id;
        vscodeMock.treeEvents = [];

        explorer.refreshAgentReviewMarks(BOOK, 'Ledger');

        const redrawn = (vscodeMock.treeEvents as XlideNode[]).map((node) => `${node.kind}:${node.kind === 'folder' ? node.folder : node.moduleName}`);
        expect(redrawn).toEqual(expect.arrayContaining([
            'module:Ledger',
            'folder:Accounts',
            'folder:Accounts.Ledger',
        ]));
        expect(redrawn).not.toContain('folder:Accounts.Billing');
        // Same render id: VS Code keeps the row's expansion as the user left it.
        expect(explorer.getTreeItem(folder('Accounts')).id).toBe(before);
    });

    it('finds the module however the agent cased its name', async () => {
        const { explorer } = await drawnTree();
        vscodeMock.treeEvents = [];

        explorer.refreshAgentReviewMarks(BOOK, 'LEDGER');

        const redrawn = (vscodeMock.treeEvents as XlideNode[]).map((node) => `${node.kind}:${node.kind === 'folder' ? node.folder : node.moduleName}`);
        expect(redrawn).toEqual(expect.arrayContaining(['module:Ledger', 'folder:Accounts.Ledger']));
    });

    it('redraws no folder in the flat layout', async () => {
        const explorer = new ProjectExplorer(fakeBridge());
        const [project] = await explorer.getChildren();
        await explorer.getChildren(project);
        vscodeMock.treeEvents = [];

        explorer.refreshAgentReviewMarks(BOOK, 'Ledger');

        expect((vscodeMock.treeEvents as XlideNode[]).every((node) => node.kind !== 'folder')).toBe(true);
    });
});

describe('a module that differs from the last commit stands out', () => {
    const marks = new Map<string, { byModule: Map<string, 'modified' | 'added'>; removed: number; head: string }>();
    const asked: string[] = [];
    const source = {
        marksFor: (filePath: string) => {
            asked.push(filePath);
            return marks.get(filePath);
        },
        onDidChange: vi.fn(),
    };

    async function drawnFlatTree() {
        const explorer = new ProjectExplorer(fakeBridge(), undefined, source);
        const [project] = await explorer.getChildren();
        await explorer.getChildren(project);
        return { explorer, project };
    }

    beforeEach(() => {
        marks.clear();
        asked.length = 0;
    });

    afterEach(() => {
        keepAgentChange(BOOK, 'Ledger');
    });

    it('asks for the marks when it draws a module, so the first draw schedules them', async () => {
        const { explorer } = await drawnFlatTree();
        explorer.getTreeItem(explorer.getModuleNode(BOOK, 'Ledger')!);
        expect(asked).toContain(BOOK);
    });

    it('decorates a modified module and an added one, and no other', async () => {
        marks.set(BOOK, { byModule: new Map([['ledger', 'modified'], ['loose', 'added']]), removed: 0, head: 'abc' });
        const { explorer } = await drawnFlatTree();

        const ledger = explorer.getTreeItem(explorer.getModuleNode(BOOK, 'Ledger')!);
        expect(parseDecorationUri(ledger.resourceUri!)).toEqual({ kind: 'module', filePath: BOOK, moduleName: 'Ledger' });
        expect(String(ledger.tooltip)).toContain('Modified since the last commit');
        // The marks colour the name through the decoration; the icon and the
        // description stay what the module type says.
        expect(ledger.description).toBe('standard');
        expect((ledger.iconPath as { color?: unknown }).color).toBeUndefined();

        const loose = explorer.getTreeItem(explorer.getModuleNode(BOOK, 'Loose')!);
        expect(String(loose.tooltip)).toContain('Not in the last commit');

        expect(explorer.getTreeItem(explorer.getModuleNode(BOOK, 'Helpers')!).resourceUri).toBeUndefined();
    });

    it('lets an agent edit awaiting review win the row', async () => {
        marks.set(BOOK, { byModule: new Map([['ledger', 'modified']]), removed: 0, head: 'abc' });
        const { explorer } = await drawnFlatTree();
        await presentAgentModuleWrite(BOOK, 'Ledger', {
            before: 'Sub Old()\r\nEnd Sub\r\n',
            beforeExisted: true,
            after: 'Sub New()\r\nEnd Sub\r\n',
        });

        const item = explorer.getTreeItem(explorer.getModuleNode(BOOK, 'Ledger')!);
        expect(item.description).toBe('standard ● agent edit');
        expect(String(item.tooltip)).toContain('An agent edited this module');
    });

    it('redraws the project row and every module row when the marks land', async () => {
        const { explorer } = await drawnFlatTree();
        const before = explorer.getTreeItem(explorer.getModuleNode(BOOK, 'Ledger')!).id;
        vscodeMock.treeEvents = [];

        explorer.refreshGitMarks(BOOK);

        const redrawn = (vscodeMock.treeEvents as XlideNode[]).map((node) => `${node.kind}:${node.moduleName ?? ''}`);
        expect(redrawn).toContain('module:Ledger');
        expect(redrawn).toContain('module:Loose');
        expect(redrawn).toContain('project:');
        expect(explorer.getTreeItem(explorer.getModuleNode(BOOK, 'Ledger')!).id).toBe(before);
    });
});
