import { vi, describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as nodePath from 'path';
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

/** Minimal stand-in - decodeModuleUri only reads uri.path */
function fakeUri(uriPath: string): VscodeType.Uri {
    return { path: uriPath, toString: () => uriPath } as VscodeType.Uri;
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
    it('falls back to provider-owned mtimes when the project file cannot be statted', () => {
        const uri = fakeUri(moduleUriPath('C:/xlide-does-not-exist/book.xlsm'));
        const firstProvider = new XlideFileSystemProvider({ call: vi.fn() } as never);
        const secondProvider = new XlideFileSystemProvider({ call: vi.fn() } as never);

        const first = firstProvider.stat(uri);
        const second = secondProvider.stat(uri);

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

        const first = provider.stat(uri);
        const second = provider.stat(uri);
        expect(second.mtime).toBe(first.mtime);

        const bytes = await provider.readFile(uri);
        const afterRead = provider.stat(uri);
        expect(Buffer.from(bytes).toString('utf-8')).toBe('Public Sub T()\nEnd Sub\n');
        expect(afterRead.mtime).toBe(first.mtime);
        expect(afterRead.size).toBe(Buffer.byteLength('Public Sub T()\nEnd Sub\n', 'utf-8'));

        provider.notifyFileChanged(uri);
        const afterChange = provider.stat(uri);
        expect(afterChange.mtime).toBeGreaterThan(first.mtime);
    });

    it('bumps mtime only after a successful write', async () => {
        const bridge = {
            call: vi.fn(async () => ({ ok: true, signatureDropped: false })),
        };
        const provider = new XlideFileSystemProvider(bridge as never);
        const uri = fakeUri('/home/user/workbook.xlsm/Module1.bas');
        const before = provider.stat(uri);

        await provider.writeFile(uri, Buffer.from('Sub Saved()\nEnd Sub\n'), {
            create: false,
            overwrite: true,
        });

        const after = provider.stat(uri);
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
        fs.rmSync(tempDir, { recursive: true, force: true });
    });

    function setProjectMtime(ms: number): void {
        fs.utimesSync(projectPath, new Date(ms), new Date(ms));
    }

    it('derives module mtimes from the backing project file', () => {
        const provider = new XlideFileSystemProvider({ call: vi.fn() } as never);
        const uri = fakeUri(moduleUriPath(projectPath));

        const stat = provider.stat(uri);

        expect(stat.mtime).toBe(t0);
        expect(stat.ctime).toBe(t0);
    });

    it('bumps every open module stat when the project changes out of band', () => {
        const provider = new XlideFileSystemProvider({ call: vi.fn() } as never);
        const uriA = fakeUri(moduleUriPath(projectPath, 'ModuleA'));
        const uriB = fakeUri(moduleUriPath(projectPath, 'ModuleB'));
        expect(provider.stat(uriA).mtime).toBe(t0);
        expect(provider.stat(uriB).mtime).toBe(t0);

        setProjectMtime(t1);

        expect(provider.stat(uriA).mtime).toBe(t1);
        expect(provider.stat(uriB).mtime).toBe(t1);
    });

    it('does not flag sibling modules when the provider itself saves the project', async () => {
        const bridge = {
            call: vi.fn(async () => {
                // The bridge saves the workbook in place
                setProjectMtime(t1);
                return { ok: true, signatureDropped: false };
            }),
        };
        const provider = new XlideFileSystemProvider(bridge as never);
        const uriA = fakeUri(moduleUriPath(projectPath, 'ModuleA'));
        const uriB = fakeUri(moduleUriPath(projectPath, 'ModuleB'));
        expect(provider.stat(uriA).mtime).toBe(t0);
        expect(provider.stat(uriB).mtime).toBe(t0);

        await provider.writeFile(uriA, Buffer.from('Sub A()\nEnd Sub\n'), {
            create: false,
            overwrite: true,
        });

        expect(provider.stat(uriA).mtime).toBe(t1);
        expect(provider.stat(uriB).mtime).toBe(t0);

        // ...but a later out-of-band change (e.g. Excel VBE edit) is seen
        setProjectMtime(t2);
        expect(provider.stat(uriB).mtime).toBe(t2);
    });

    it('adopts the new project mtime via notifyFileChanged without disturbing siblings', () => {
        const provider = new XlideFileSystemProvider({ call: vi.fn() } as never);
        const uriA = fakeUri(moduleUriPath(projectPath, 'ModuleA'));
        const uriB = fakeUri(moduleUriPath(projectPath, 'ModuleB'));
        expect(provider.stat(uriA).mtime).toBe(t0);
        expect(provider.stat(uriB).mtime).toBe(t0);

        setProjectMtime(t1);
        provider.notifyFileChanged(uriA);

        expect(provider.stat(uriA).mtime).toBe(t1);
        expect(provider.stat(uriB).mtime).toBe(t0);
    });
});
