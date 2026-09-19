import { describe, expect, it, vi } from 'vitest';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());
// Each write passes straight through: the real coordinator talks to whatever
// Office application is running.
vi.mock('../src/officeWriteCoordinator', () => ({
    runWriteWithHostCoordination: vi.fn((_filePath: string, write: () => Promise<unknown>) => write()),
}));

import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import * as vscode from 'vscode';
import {
    deleteProjectModule,
    refreshProjectStateOnOutsideChange,
    renameProjectModule,
    writeProjectFormDesigner,
} from '../src/projectModuleOperations';
import { checkProjectFile } from '../src/projectFileChanges';
import { encodeFormMarkupUri, encodeModuleUri } from '../src/xlideFileSystem';

const PROJECT = process.platform === 'win32' ? 'C:\\work\\Book.xlsm' : '/work/Book.xlsm';

describe('renameProjectModule', () => {
    it('moves an editor on the module to its new name, and leaves one with unsaved edits', async () => {
        // Left on the old name, an editor showed a module that no longer
        // existed, and its text kept the old module alive for analysis.
        const oldUri = encodeModuleUri(PROJECT, 'OldName');
        const clean = { input: new vscode.TabInputText(oldUri), isDirty: false, isPreview: false, group: { viewColumn: 2 } };
        const dirty = { input: new vscode.TabInputText(oldUri), isDirty: true, isPreview: false, group: { viewColumn: 1 } };
        const window = vscode.window as unknown as {
            tabGroups: { all: unknown[]; close: ReturnType<typeof vi.fn> };
            showTextDocument: ReturnType<typeof vi.fn>;
            visibleTextEditors: unknown[];
        };
        window.tabGroups.all = [{ tabs: [dirty] }, { tabs: [clean] }];
        window.tabGroups.close = vi.fn(async () => true);
        window.showTextDocument = vi.fn(async () => undefined);
        window.visibleTextEditors = [];
        const deps = {
            bridge: { call: vi.fn(async () => ({ ok: true, signatureDropped: false })) },
            explorer: { refresh: vi.fn() },
            fsProvider: { notifyFileChanged: vi.fn() },
            vbaIndex: { invalidate: vi.fn() },
        };

        await renameProjectModule(deps as never, { filePath: PROJECT, moduleName: 'OldName', newName: 'NewName' });

        expect(window.showTextDocument).toHaveBeenCalledTimes(1);
        const [target, options] = window.showTextDocument.mock.calls[0] as [{ path: string }, { viewColumn: number }];
        expect(target.path).toMatch(/\/NewName\.bas$/);
        expect(options.viewColumn).toBe(2);
        expect(window.tabGroups.close).toHaveBeenCalledWith(clean, true);
        expect(window.tabGroups.close).not.toHaveBeenCalledWith(dirty, expect.anything());
        window.tabGroups.all = [];
    });

    it('follows to the module the engine made, when that is not the name asked for', async () => {
        // An Access form's module keeps its prefix: renamed to `Customers`,
        // Form_Calculator became Form_Customers, and the editor was moved to
        // a `Customers` that did not exist.
        const oldUri = encodeModuleUri(PROJECT, 'Form_Calculator');
        const tab = { input: new vscode.TabInputText(oldUri), isDirty: false, isPreview: false, group: { viewColumn: 1 } };
        const window = vscode.window as unknown as {
            tabGroups: { all: unknown[]; close: ReturnType<typeof vi.fn> };
            showTextDocument: ReturnType<typeof vi.fn>;
            visibleTextEditors: unknown[];
        };
        window.tabGroups.all = [{ tabs: [tab] }];
        window.tabGroups.close = vi.fn(async () => true);
        window.showTextDocument = vi.fn(async () => undefined);
        window.visibleTextEditors = [];
        const deps = {
            bridge: { call: vi.fn(async () => ({ ok: true, signatureDropped: false, moduleName: 'Form_Customers' })) },
            explorer: { refresh: vi.fn() },
            fsProvider: { notifyFileChanged: vi.fn() },
            vbaIndex: { invalidate: vi.fn() },
        };

        const result = await renameProjectModule(deps as never, { filePath: PROJECT, moduleName: 'Form_Calculator', newName: 'Customers' });

        expect(result.moduleName).toBe('Form_Customers');
        const [target] = window.showTextDocument.mock.calls[0] as [{ path: string }];
        expect(target.path).toMatch(/\/Form_Customers\.bas$/);
        window.tabGroups.all = [];
    });

    it("moves a form's markup and designer too, the markup still as XML", async () => {
        // Only the code followed: the designer and the markup went on showing
        // a form that no longer existed.
        const oldMarkup = encodeFormMarkupUri(PROJECT, 'OldForm');
        const markupTab = { input: new vscode.TabInputText(oldMarkup), isDirty: false, isPreview: false, group: { viewColumn: 1 } };
        const designerTab = {
            input: new vscode.TabInputCustom(oldMarkup, 'xlideFormDesigner'),
            isDirty: false,
            isPreview: true,
            group: { viewColumn: 2 },
        };
        const shownMarkup = { uri: encodeFormMarkupUri(PROJECT, 'NewForm'), languageId: 'plaintext' };
        const mocked = vscode as unknown as {
            window: {
                tabGroups: { all: unknown[]; close: ReturnType<typeof vi.fn> };
                showTextDocument: ReturnType<typeof vi.fn>;
                visibleTextEditors: unknown[];
            };
            workspace: { textDocuments: unknown[] };
            commands: { executeCommand: ReturnType<typeof vi.fn> };
            languages: { setTextDocumentLanguage: ReturnType<typeof vi.fn> };
        };
        mocked.window.tabGroups.all = [{ tabs: [markupTab] }, { tabs: [designerTab] }];
        mocked.window.tabGroups.close = vi.fn(async () => true);
        mocked.window.showTextDocument = vi.fn(async () => ({ document: shownMarkup }));
        mocked.window.visibleTextEditors = [];
        mocked.workspace.textDocuments = [{ uri: oldMarkup, languageId: 'xml' }];
        mocked.commands.executeCommand = vi.fn(async () => undefined);
        mocked.languages.setTextDocumentLanguage = vi.fn(async (document: unknown) => document);
        const deps = {
            bridge: { call: vi.fn(async () => ({ ok: true, signatureDropped: false })) },
            explorer: { refresh: vi.fn() },
            fsProvider: { notifyFileChanged: vi.fn() },
            vbaIndex: { invalidate: vi.fn() },
        };

        await renameProjectModule(deps as never, { filePath: PROJECT, moduleName: 'OldForm', newName: 'NewForm' });

        const [shown] = mocked.window.showTextDocument.mock.calls[0] as [{ path: string }];
        expect(shown.path).toMatch(/\/NewForm\.form$/);
        expect(mocked.languages.setTextDocumentLanguage).toHaveBeenCalledWith(shownMarkup, 'xml');
        expect(mocked.commands.executeCommand).toHaveBeenCalledWith(
            'vscode.openWith',
            expect.objectContaining({ path: expect.stringMatching(/\/NewForm\.form$/) }),
            'xlideFormDesigner',
            { viewColumn: 2, preview: true, preserveFocus: true },
        );
        expect(mocked.window.tabGroups.close).toHaveBeenCalledWith(markupTab, true);
        expect(mocked.window.tabGroups.close).toHaveBeenCalledWith(designerTab, true);
        expect(deps.fsProvider.notifyFileChanged.mock.calls.map(([uri]) => (uri as { path: string }).path))
            .toEqual([expect.stringMatching(/\/OldForm\.bas$/), expect.stringMatching(/\/OldForm\.form$/)]);
        mocked.window.tabGroups.all = [];
        mocked.workspace.textDocuments = [];
    });
});

describe('deleteProjectModule', () => {
    it("closes a form's code, markup and designer, and leaves other editors", async () => {
        const code = { input: new vscode.TabInputText(encodeModuleUri(PROJECT, 'Gone')) };
        const markup = { input: new vscode.TabInputText(encodeFormMarkupUri(PROJECT, 'Gone')) };
        const designer = { input: new vscode.TabInputCustom(encodeFormMarkupUri(PROJECT, 'Gone'), 'xlideFormDesigner') };
        const other = { input: new vscode.TabInputText(encodeModuleUri(PROJECT, 'Kept')) };
        const window = vscode.window as unknown as { tabGroups: { all: unknown[]; close: ReturnType<typeof vi.fn> } };
        window.tabGroups.all = [{ tabs: [code, markup, other] }, { tabs: [designer] }];
        window.tabGroups.close = vi.fn(async () => true);
        const deps = {
            bridge: { call: vi.fn(async () => ({ ok: true, signatureDropped: false })) },
            explorer: { refresh: vi.fn() },
            fsProvider: { notifyFileChanged: vi.fn() },
            vbaIndex: { invalidate: vi.fn() },
        };

        await deleteProjectModule(deps as never, { filePath: PROJECT, moduleName: 'Gone' });

        expect(window.tabGroups.close.mock.calls.map(([tab]) => tab)).toEqual([code, markup, designer]);
        window.tabGroups.all = [];
    });
});

describe('refreshProjectStateOnOutsideChange', () => {
    it('refreshes only when an outside change reached the VBA', async () => {
        const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-outside-refresh-'));
        const book = path.join(dir, 'Book.xlsm');
        let modules = [{ name: 'Module1', source: 'Sub A()\r\nEnd Sub\r\n' }];
        const indexChanged = new vscode.EventEmitter<{ projectPath: string; moduleName?: string }>();
        const deps = {
            bridge: { call: vi.fn(async () => modules) },
            explorer: { refresh: vi.fn() },
            vbaIndex: {
                onDidChange: indexChanged.event,
                // As the real index does: dropping a project announces it.
                invalidate: vi.fn((projectPath: string) => indexChanged.fire({ projectPath })),
            },
        };
        const subscription = refreshProjectStateOnOutsideChange(deps as never);
        let saves = 0;
        const saveOutsideXlide = async (): Promise<void> => {
            saves += 1;
            fs.writeFileSync(book, `saved ${saves}`.padEnd(10 + saves, '.'));
            checkProjectFile(book);
            await new Promise((resolve) => setTimeout(resolve, 0));
        };
        try {
            fs.writeFileSync(book, 'first');
            checkProjectFile(book);

            // The first outside change has nothing to compare with.
            await saveOutsideXlide();
            expect(deps.explorer.refresh).toHaveBeenCalledTimes(1);

            // Cells saved again, AutoSave, Access's own writes: the VBA is the same.
            await saveOutsideXlide();
            await saveOutsideXlide();
            expect(deps.explorer.refresh).toHaveBeenCalledTimes(1);

            // The VBE saved a change to a module.
            modules = [{ name: 'Module1', source: 'Sub A()\r\n    x = 1\r\nEnd Sub\r\n' }];
            await saveOutsideXlide();
            expect(deps.explorer.refresh).toHaveBeenCalledTimes(2);

            // XLIDE wrote a module, so the project's cached state moved on. An
            // outside change is then compared from scratch, even one that puts
            // the VBA back where the last outside change left it.
            indexChanged.fire({ projectPath: book, moduleName: 'Module1' });
            await saveOutsideXlide();
            expect(deps.explorer.refresh).toHaveBeenCalledTimes(3);
        } finally {
            subscription.dispose();
            fs.rmSync(dir, { recursive: true, force: true });
        }
    });
});

describe('writeProjectFormDesigner', () => {
    it('announces the write to the form markup document, which is what shows the designer', async () => {
        // It announced the code document instead, which a designer write does
        // not change, so an open markup document kept the old designer, and
        // saving it later wrote that designer back over the import.
        const notifyFileChanged = vi.fn();
        const deps = {
            bridge: { call: vi.fn(async () => ({ ok: true, signatureDropped: false })) },
            explorer: { refresh: vi.fn() },
            fsProvider: { notifyFileChanged },
            vbaIndex: { invalidate: vi.fn() },
        };

        await writeProjectFormDesigner(deps as never, { filePath: PROJECT, moduleName: 'UserForm1', frx: Buffer.from('frx') });

        expect(notifyFileChanged).toHaveBeenCalledTimes(1);
        expect(String(notifyFileChanged.mock.calls[0][0].path)).toMatch(/\/UserForm1\.form$/);
    });
});
