import { vi, describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as nodePath from 'path';
import * as vscode from 'vscode';
import type * as VscodeType from 'vscode';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());

import {
    decodeModuleUri,
    encodeModuleUri,
    moduleIdentityKey,
    sameProjectPath,
    projectIdentityKey,
    XlideFileSystemProvider,
} from '../src/xlideFileSystem';
import { recordProjectWrite } from '../src/projectFileChanges';

/** Minimal stand-in - decodeModuleUri only reads uri.path */
function fakeUri(uriPath: string): VscodeType.Uri {
    return { scheme: 'xlide-vba', path: uriPath, toString: () => uriPath } as VscodeType.Uri;
}

function moduleUriPath(projectPath: string, moduleName = 'Module1'): string {
    const forward = projectPath.replace(/\\/g, '/');
    const base = forward.startsWith('/') ? forward : `/${forward}`;
    return `${base}/${moduleName}.bas`;
}

describe('decodeModuleUri', () => {
    it('decodes modules from every macro container, not only Excel', () => {
        const cases: Array<[string, string]> = [
            ['/work/Report.docm/ThisDocument.bas', 'ThisDocument'],
            ['/work/Letters.dotm/Module1.bas', 'Module1'],
            ['/work/Legacy.doc/CGreeter.bas', 'CGreeter'],
            ['/work/Deck.pptm/CDeck.bas', 'CDeck'],
            ['/work/Show.ppt/Module1.bas', 'Module1'],
            ['/work/Book.xls/Module1.bas', 'Module1'],
            ['/work/Data.accdb/CAudit.bas', 'CAudit'],
            ['/work/Old.mdb/Module1.bas', 'Module1'],
        ];
        for (const [uriPath, expectedModule] of cases) {
            const { moduleName, projectPath } = decodeModuleUri(fakeUri(uriPath));
            expect(moduleName, uriPath).toBe(expectedModule);
            expect(projectPath.toLowerCase(), uriPath).toContain(uriPath.split('/')[2].toLowerCase());
        }
    });

    it('decodes the module name from a basic path', () => {
        const { moduleName } = decodeModuleUri(fakeUri('/home/user/workbook.xlsm/Module1.bas'));
        expect(moduleName).toBe('Module1');
    });

    it('URL-decodes spaces in module names', () => {
        const { moduleName } = decodeModuleUri(
            fakeUri('/home/user/workbook.xlsm/My%20Module.bas'),
        );
        expect(moduleName).toBe('My Module');
    });

    it('works with .xlsb extension', () => {
        const { moduleName } = decodeModuleUri(fakeUri('/home/user/book.xlsb/Sheet1.bas'));
        expect(moduleName).toBe('Sheet1');
    });

    it('works with .xlam extension', () => {
        const { moduleName } = decodeModuleUri(fakeUri('/home/user/addin.xlam/Helpers.bas'));
        expect(moduleName).toBe('Helpers');
    });

    it('projectPath ends with the project filename', () => {
        const { projectPath } = decodeModuleUri(fakeUri('/home/user/mybook.xlsm/Module1.bas'));
        expect(projectPath.endsWith('mybook.xlsm')).toBe(true);
    });

    it('throws on a path with no recognised project extension', () => {
        expect(() => decodeModuleUri(fakeUri('/home/user/file.txt'))).toThrow();
    });

    it('throws on a path with missing module segment', () => {
        expect(() => decodeModuleUri(fakeUri('/home/user/workbook.xlsm/'))).toThrow();
    });

    it('round-trips project paths containing reserved characters (# and %)', () => {
        // Built from a structured path (Uri.from), these must survive encode ->
        // decode rather than being split on '#' or having '%xx' decoded.
        for (const projectPath of [
            '/home/user/My #1 Book.xlsm',
            '/home/user/50%done.xlsm',
            '/home/user/report%20v2.xlsm',
        ]) {
            const uri = encodeModuleUri(projectPath, 'Module1');
            const decoded = decodeModuleUri(uri);
            expect(decoded.moduleName).toBe('Module1');
            expect(decoded.projectPath.replace(/\\/g, '/')).toBe(projectPath);
        }
    });
});

describe('project identity helpers', () => {
    it('normalizes project paths case-insensitively on Windows only', () => {
        expect(projectIdentityKey('C:/Repo/Book.xlsm', 'win32')).toBe('c:\\repo\\book.xlsm');
        expect(projectIdentityKey('/Users/me/Book.xlsm', 'darwin')).toBe('/Users/me/Book.xlsm');
        expect(sameProjectPath('C:/Repo/Book.xlsm', 'c:/repo/book.xlsm', 'win32')).toBe(true);
        expect(sameProjectPath('C:/Repo/Book.xlsm', 'C:\\Repo\\Book.xlsm', 'win32')).toBe(true);
        expect(sameProjectPath('/repo/Book.xlsm', '/repo/book.xlsm', 'linux')).toBe(false);
    });

    it('normalizes VBA module identity independent of project identity', () => {
        expect(moduleIdentityKey('Module1')).toBe('module1');
        expect(moduleIdentityKey('Person')).toBe(moduleIdentityKey('person'));
    });
});

describe('XlideFileSystemProvider stats', () => {
    it('falls back to provider-owned mtimes when the project file cannot be statted', async () => {
        const uri = fakeUri(moduleUriPath('C:/xlide-does-not-exist/book.xlsm'));
        const firstProvider = new XlideFileSystemProvider({ call: vi.fn() } as never);
        const secondProvider = new XlideFileSystemProvider({ call: vi.fn() } as never);

        const first = await firstProvider.stat(uri);
        const second = await secondProvider.stat(uri);

        expect(first.mtime).toBeGreaterThan(0);
        expect(first.ctime).toBe(first.mtime);
        expect(second.mtime).toBeGreaterThan(0);
        expect(second.ctime).toBe(second.mtime);
    });

    it('keeps mtime stable across stat/read checks until an explicit change event', async () => {
        const bridge = {
            call: vi.fn(async () => ({ source: 'Public Sub T()\nEnd Sub\n' })),
        };
        const provider = new XlideFileSystemProvider(bridge as never);
        const uri = fakeUri('/home/user/workbook.xlsm/Module1.bas');

        const first = await provider.stat(uri);
        const second = await provider.stat(uri);
        expect(second.mtime).toBe(first.mtime);

        const bytes = await provider.readFile(uri);
        const afterRead = await provider.stat(uri);
        expect(Buffer.from(bytes).toString('utf-8')).toBe('Public Sub T()\nEnd Sub\n');
        expect(afterRead.mtime).toBe(first.mtime);
        expect(afterRead.size).toBe(Buffer.byteLength('Public Sub T()\nEnd Sub\n', 'utf-8'));

        provider.notifyFileChanged(uri);
        const afterChange = await provider.stat(uri);
        expect(afterChange.mtime).toBeGreaterThan(first.mtime);
    });

    it('bumps mtime only after a successful write', async () => {
        const bridge = {
            call: vi.fn(async () => ({ ok: true, signatureDropped: false })),
        };
        const provider = new XlideFileSystemProvider(bridge as never);
        const uri = fakeUri('/home/user/workbook.xlsm/Module1.bas');
        const before = await provider.stat(uri);

        await provider.writeFile(uri, Buffer.from('Sub Saved()\nEnd Sub\n'), {
            create: false,
            overwrite: true,
        });

        const after = await provider.stat(uri);
        expect(after.mtime).toBeGreaterThan(before.mtime);
        expect(after.size).toBe(Buffer.byteLength('Sub Saved()\nEnd Sub\n', 'utf-8'));
    });
});

describe('XlideFileSystemProvider stats (real project file)', () => {
    const t0 = Date.parse('2024-01-01T00:00:00Z');
    const t1 = Date.parse('2024-01-02T00:00:00Z');
    const t2 = Date.parse('2024-01-03T00:00:00Z');
    let tempDir: string;
    let projectPath: string;

    beforeEach(() => {
        tempDir = fs.mkdtempSync(nodePath.join(os.tmpdir(), 'xlide-fs-stats-'));
        projectPath = nodePath.join(tempDir, 'book.xlsm');
        fs.writeFileSync(projectPath, 'stub');
        setProjectMtime(t0);
    });

    afterEach(() => {
        openDocuments.length = 0;
        fs.rmSync(tempDir, { recursive: true, force: true });
    });

    function setProjectMtime(ms: number): void {
        fs.utimesSync(projectPath, new Date(ms), new Date(ms));
    }

    interface FakeDocument {
        uri: VscodeType.Uri;
        isClosed: boolean;
        isDirty: boolean;
        text: string;
        getText(): string;
    }

    /** The documents VS Code has open, which is what the provider compares. */
    const openDocuments = vscode.workspace.textDocuments as unknown as FakeDocument[];

    /**
     * An engine over a module map; a write saves the file, as the real one
     * does, storing the text the way `store` says.
     */
    function fakeEngine(modules: Record<string, string>, savedAt: number, store = (source: string) => source) {
        return {
            call: vi.fn(async (method: string, params: { module: string; source?: string }) => {
                if (method === 'readModule') {
                    if (!(params.module in modules)) {
                        throw new Error(`Module "${params.module}" does not exist.`);
                    }
                    return { source: modules[params.module] };
                }
                if (method === 'writeModule') {
                    modules[params.module] = store(params.source ?? '');
                    setProjectMtime(savedAt);
                    return { ok: true, signatureDropped: false };
                }
                throw new Error(`unexpected engine call ${method}`);
            }),
        };
    }

    /** Opens the modules the way VS Code does: stat, then read, then a document. */
    async function openModules(provider: XlideFileSystemProvider, ...uris: VscodeType.Uri[]): Promise<FakeDocument[]> {
        const documents: FakeDocument[] = [];
        for (const uri of uris) {
            await provider.stat(uri);
            const text = Buffer.from(await provider.readFile(uri)).toString('utf-8');
            const document: FakeDocument = { uri, isClosed: false, isDirty: false, text, getText() { return this.text; } };
            openDocuments.push(document);
            documents.push(document);
        }
        return documents;
    }

    /** A save from the editor: the document holds what it wrote, clean. */
    async function save(provider: XlideFileSystemProvider, document: FakeDocument, text: string): Promise<void> {
        await provider.writeFile(document.uri, Buffer.from(text), { create: false, overwrite: true });
        document.text = text;
        document.isDirty = false;
    }

    function changedUris(provider: XlideFileSystemProvider): string[] {
        const fired: string[] = [];
        provider.onDidChangeFile((events) => fired.push(...events.map((event) => event.uri.toString())));
        return fired;
    }

    it('answers a read of a module the project no longer has as a missing file', async () => {
        // A document VS Code still held for a deleted or renamed module was
        // read again, and the engine's own error was logged as the provider
        // failing; FileNotFound it takes as the file being gone.
        const provider = new XlideFileSystemProvider({
            call: vi.fn(async (_method: string, params: { module: string }) => {
                throw new Error(`Module not found: ${params.module}`);
            }),
        } as never);
        const uri = fakeUri(moduleUriPath(projectPath, 'Gone'));
        const fileNotFound = vi.spyOn(vscode.FileSystemError, 'FileNotFound');
        try {
            await expect(provider.readFile(uri)).rejects.toThrow();
            expect(fileNotFound).toHaveBeenCalledWith(uri);
        } finally {
            fileNotFound.mockRestore();
        }
    });

    it('derives module mtimes from the backing project file', async () => {
        const provider = new XlideFileSystemProvider({ call: vi.fn() } as never);
        const uri = fakeUri(moduleUriPath(projectPath));

        const stat = await provider.stat(uri);

        expect(stat.mtime).toBe(t0);
        expect(stat.ctime).toBe(t0);
    });

    it('moves forward only the open modules a change made outside XLIDE reached', async () => {
        const modules = { ModuleA: 'Sub A()\r\nEnd Sub\r\n', ModuleB: 'Sub B()\r\nEnd Sub\r\n' };
        const provider = new XlideFileSystemProvider(fakeEngine(modules, t1) as never);
        const uriA = fakeUri(moduleUriPath(projectPath, 'ModuleA'));
        const uriB = fakeUri(moduleUriPath(projectPath, 'ModuleB'));
        await openModules(provider, uriA, uriB);
        const fired = changedUris(provider);

        // The VBE saves the workbook with ModuleA edited.
        modules.ModuleA = 'Sub A()\r\n    x = 1\r\nEnd Sub\r\n';
        setProjectMtime(t1);

        // ModuleA is reported changed, so VS Code reloads it; ModuleB, which
        // the change did not reach, keeps its mtime.
        expect((await provider.stat(uriB)).mtime).toBe(t0);
        expect((await provider.stat(uriA)).mtime).toBe(t1);
        expect(fired).toEqual([uriA.toString()]);
    });

    it('holds unsaved edits to a conflict only where the change reached their module', async () => {
        const modules = { ModuleA: 'Sub A()\r\nEnd Sub\r\n', ModuleB: 'Sub B()\r\nEnd Sub\r\n' };
        const provider = new XlideFileSystemProvider(fakeEngine(modules, t1) as never);
        const uriA = fakeUri(moduleUriPath(projectPath, 'ModuleA'));
        const uriB = fakeUri(moduleUriPath(projectPath, 'ModuleB'));
        const [editedA, editedB] = await openModules(provider, uriA, uriB);
        const before = await provider.stat(uriB);
        for (const document of [editedA, editedB]) {
            document.text += "' typed\r\n";
            document.isDirty = true;
        }

        modules.ModuleB = 'Sub B()\r\n    x = 1\r\nEnd Sub\r\n';
        setProjectMtime(t1);

        // VS Code reports "File Modified Since" on a newer mtime AND a
        // different size, so both have to move for ModuleB's save to stop.
        const afterB = await provider.stat(uriB);
        expect(afterB.mtime).toBe(t1);
        expect(afterB.size).not.toBe(before.size);
        expect(afterB.size).toBe(Buffer.byteLength(modules.ModuleB, 'utf-8'));
        expect((await provider.stat(uriA)).mtime).toBe(t0);
    });

    it('raises no conflict over its own save the engine stored differently, when the change was elsewhere', async () => {
        // The engine drops blank lines above a module's code. The save was
        // recorded as the text written, so the next change anywhere in the
        // project looked like a change to this module and stopped its save.
        const modules = { ModuleA: 'Sub A()\r\nEnd Sub\r\n', ModuleB: 'Sub B()\r\nEnd Sub\r\n' };
        const provider = new XlideFileSystemProvider(
            fakeEngine(modules, t1, (source) => source.replace(/^(?:\r\n)+/, '')) as never,
        );
        const uriA = fakeUri(moduleUriPath(projectPath, 'ModuleA'));
        const uriB = fakeUri(moduleUriPath(projectPath, 'ModuleB'));
        const [documentA] = await openModules(provider, uriA, uriB);
        await save(provider, documentA, '\r\n\r\nSub A()\r\nEnd Sub\r\n');
        const saved = await provider.stat(uriA);
        documentA.text += "' typed\r\n";
        documentA.isDirty = true;

        modules.ModuleB = 'Sub B()\r\n    x = 1\r\nEnd Sub\r\n';
        setProjectMtime(t2);

        const after = await provider.stat(uriA);
        expect(after.mtime).toBe(saved.mtime);
        expect(after.size).toBe(saved.size);
    });

    it('holds a save to a conflict over a change that kept the byte length', async () => {
        // VS Code alone would let this save through: its check needs the size
        // to differ. XLIDE compared the content, so it knows better.
        const modules = { ModuleA: 'Sub Mine()\r\nEnd Sub\r\n' };
        const provider = new XlideFileSystemProvider(fakeEngine(modules, t1) as never);
        const uriA = fakeUri(moduleUriPath(projectPath, 'ModuleA'));
        const [documentA] = await openModules(provider, uriA);
        const given = await provider.stat(uriA);
        documentA.text += "' typed\r\n";
        documentA.isDirty = true;

        modules.ModuleA = 'Sub Them()\r\nEnd Sub\r\n';
        setProjectMtime(t1);

        const after = await provider.stat(uriA);
        expect(after.mtime).toBeGreaterThan(given.mtime);
        expect(after.size).not.toBe(given.size);
    });

    it('holds a save to a conflict when the changed module cannot be read', async () => {
        // Deleted in the VBE: saving the editor would have written the module
        // back into the workbook without a prompt.
        const modules: Record<string, string> = { ModuleA: 'Sub A()\r\nEnd Sub\r\n' };
        const provider = new XlideFileSystemProvider(fakeEngine(modules, t1) as never);
        const uriA = fakeUri(moduleUriPath(projectPath, 'ModuleA'));
        const [documentA] = await openModules(provider, uriA);
        const given = await provider.stat(uriA);
        documentA.text += "' typed\r\n";
        documentA.isDirty = true;

        delete modules.ModuleA;
        setProjectMtime(t1);

        const after = await provider.stat(uriA);
        expect(after.mtime).toBeGreaterThan(given.mtime);
        expect(after.size).not.toBe(given.size);
    });

    it('moves nothing when the file changed and none of its open modules did', async () => {
        // Cells written, a sheet added: the workbook is newer, every module the same.
        const modules = { ModuleA: 'Sub A()\r\nEnd Sub\r\n', ModuleB: 'Sub B()\r\nEnd Sub\r\n' };
        const provider = new XlideFileSystemProvider(fakeEngine(modules, t1) as never);
        const uriA = fakeUri(moduleUriPath(projectPath, 'ModuleA'));
        const uriB = fakeUri(moduleUriPath(projectPath, 'ModuleB'));
        await openModules(provider, uriA, uriB);
        const fired = changedUris(provider);

        setProjectMtime(t1);

        expect((await provider.stat(uriA)).mtime).toBe(t0);
        expect((await provider.stat(uriB)).mtime).toBe(t0);
        expect(fired).toEqual([]);
    });

    it('keeps what a module was just given when the sweep of closed documents lands while it opens', async () => {
        // The sweep a second after a tab closes dropped the record of every
        // module VS Code did not list as open - including one read a moment
        // before, whose document VS Code had not listed yet. With the record
        // gone, the next change anywhere in the file stopped its save at a
        // conflict nobody caused (the integration suite hit it 2 ms apart).
        vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'Date'] });
        try {
            const modules = { ModuleA: 'Sub A()\r\nEnd Sub\r\n', ModuleB: 'Sub B()\r\nEnd Sub\r\n' };
            const provider = new XlideFileSystemProvider(fakeEngine(modules, t1) as never);
            const onClose = vi.mocked(vscode.workspace.onDidCloseTextDocument).mock.calls.at(-1)![0] as (doc: unknown) => void;
            const uriA = fakeUri(moduleUriPath(projectPath, 'ModuleA'));

            onClose({ uri: fakeUri(moduleUriPath(projectPath, 'Closed')) });
            vi.advanceTimersByTime(999);
            // VS Code reads the module, then lists its document.
            const before = await provider.stat(uriA);
            const text = Buffer.from(await provider.readFile(uriA)).toString('utf-8');
            vi.advanceTimersByTime(1);
            openDocuments.push({ uri: uriA, isClosed: false, isDirty: true, text: `${text}' typed\r\n`, getText() { return this.text; } });

            modules.ModuleB = 'Sub B()\r\n    x = 1\r\nEnd Sub\r\n';
            setProjectMtime(t1);

            const after = await provider.stat(uriA);
            expect(after.mtime).toBe(before.mtime);
            expect(after.size).toBe(Buffer.byteLength(modules.ModuleA, 'utf-8'));
        } finally {
            vi.useRealTimers();
        }
    });

    it('forgets a module read but never opened at a later sweep', async () => {
        vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'Date'] });
        try {
            const provider = new XlideFileSystemProvider(fakeEngine({ ModuleA: 'Sub A()\r\nEnd Sub\r\n' }, t1) as never);
            const onClose = vi.mocked(vscode.workspace.onDidCloseTextDocument).mock.calls.at(-1)![0] as (doc: unknown) => void;
            const uriA = fakeUri(moduleUriPath(projectPath, 'ModuleA'));

            onClose({ uri: fakeUri(moduleUriPath(projectPath, 'Closed')) });
            vi.advanceTimersByTime(500);
            await provider.stat(uriA);
            await provider.readFile(uriA);
            vi.advanceTimersByTime(500);
            // Kept through the first sweep: an entry this new keeps its mtime.
            setProjectMtime(t1);
            expect((await provider.stat(uriA)).mtime).toBe(t0);

            vi.advanceTimersByTime(1000);
            // Gone at the next: a fresh entry takes the file's mtime.
            expect((await provider.stat(uriA)).mtime).toBe(t1);
        } finally {
            vi.useRealTimers();
        }
    });

    it('does not flag sibling modules when the provider itself saves the project', async () => {
        const modules = { ModuleA: 'Sub A()\r\nEnd Sub\r\n', ModuleB: 'Sub B()\r\nEnd Sub\r\n' };
        const provider = new XlideFileSystemProvider(fakeEngine(modules, t1) as never);
        const uriA = fakeUri(moduleUriPath(projectPath, 'ModuleA'));
        const uriB = fakeUri(moduleUriPath(projectPath, 'ModuleB'));
        const [documentA] = await openModules(provider, uriA, uriB);

        await save(provider, documentA, 'Sub A()\r\n    y = 2\r\nEnd Sub\r\n');

        expect((await provider.stat(uriA)).mtime).toBe(t1);
        expect((await provider.stat(uriB)).mtime).toBe(t0);

        // ...but a later change made outside XLIDE (a VBE edit) is seen
        modules.ModuleB = 'Sub B()\r\n    z = 3\r\nEnd Sub\r\n';
        setProjectMtime(t2);
        expect((await provider.stat(uriB)).mtime).toBe(t2);
    });

    it('still moves a module changed outside XLIDE when an XLIDE write lands before anything looked', async () => {
        // The VBE saves ModuleB, then a save of ModuleA runs before any stat.
        // Taking the stamp that save leaves as the only change hid ModuleB's
        // edit, and a later save of ModuleB overwrote it without a conflict.
        const modules = { ModuleA: 'Sub A()\r\nEnd Sub\r\n', ModuleB: 'Sub B()\r\nEnd Sub\r\n' };
        const provider = new XlideFileSystemProvider(fakeEngine(modules, t2) as never);
        const uriA = fakeUri(moduleUriPath(projectPath, 'ModuleA'));
        const uriB = fakeUri(moduleUriPath(projectPath, 'ModuleB'));
        const [documentA] = await openModules(provider, uriA, uriB);

        modules.ModuleB = 'Sub B()\r\n    z = 3\r\nEnd Sub\r\n';
        setProjectMtime(t1);
        await save(provider, documentA, 'Sub A()\r\n    y = 2\r\nEnd Sub\r\n');

        expect((await provider.stat(uriB)).mtime).toBe(t2);
        expect((await provider.stat(uriA)).mtime).toBe(t2);
    });

    it('adopts the new project mtime via notifyFileChanged without disturbing siblings', async () => {
        const modules = { ModuleA: 'Sub A()\r\nEnd Sub\r\n', ModuleB: 'Sub B()\r\nEnd Sub\r\n' };
        const provider = new XlideFileSystemProvider(fakeEngine(modules, t1) as never);
        const uriA = fakeUri(moduleUriPath(projectPath, 'ModuleA'));
        const uriB = fakeUri(moduleUriPath(projectPath, 'ModuleB'));
        await openModules(provider, uriA, uriB);

        // An agent tool writes ModuleA through the engine, the way every XLIDE
        // write runs, and then tells the provider.
        await recordProjectWrite(projectPath, async () => {
            modules.ModuleA = 'Sub A()\r\n    y = 2\r\nEnd Sub\r\n';
            setProjectMtime(t1);
        });
        provider.notifyFileChanged(uriA);

        expect((await provider.stat(uriA)).mtime).toBe(t1);
        expect((await provider.stat(uriB)).mtime).toBe(t0);
    });

    it('holds unsaved edits to a conflict after an XLIDE write to their module', async () => {
        // The mtime moved and the size did not, so VS Code let the editor's
        // save overwrite the agent's write without a prompt.
        const modules = { ModuleA: 'Sub A()\r\nEnd Sub\r\n' };
        const provider = new XlideFileSystemProvider(fakeEngine(modules, t1) as never);
        const uriA = fakeUri(moduleUriPath(projectPath, 'ModuleA'));
        await openModules(provider, uriA);
        const given = await provider.stat(uriA);

        await recordProjectWrite(projectPath, async () => {
            modules.ModuleA = 'Sub A()\r\n    agent = 1\r\nEnd Sub\r\n';
            setProjectMtime(t1);
        });
        provider.notifyFileChanged(uriA);

        const after = await provider.stat(uriA);
        expect(after.mtime).toBeGreaterThan(given.mtime);
        expect(after.size).not.toBe(given.size);
    });
});
