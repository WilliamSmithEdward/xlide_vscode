import { vi, describe, it, expect } from 'vitest';
import type * as VscodeType from 'vscode';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

vi.mock('vscode', () => {
    class Disposable {
        constructor(private readonly fn: () => void = () => undefined) {}
        dispose(): void { this.fn(); }
    }
    class EventEmitter<T> {
        readonly events: T[] = [];
        readonly event = () => new Disposable();
        fire(value: T): void { this.events.push(value); }
        dispose(): void { /* no-op */ }
    }
    return {
        Disposable,
        EventEmitter,
        FileType: { File: 1 },
        FileChangeType: { Changed: 1 },
        FileSystemError: {
            NoPermissions: (message: string) => new Error(message),
            Unavailable: (message: string) => new Error(message),
        },
        window: {
            showWarningMessage: vi.fn(() => Promise.resolve(undefined)),
        },
        commands: {
            executeCommand: vi.fn(),
        },
        Uri: {
            file: (fsPath: string) => ({ fsPath, path: fsPath, toString: () => fsPath }),
            parse: (value: string) => ({ path: value, toString: () => value }),
        },
    };
});
vi.mock('../src/pythonBridge', () => ({ PythonBridge: class PythonBridge {} }));
vi.mock('../src/liveShare', () => ({
    decodeRemoteModuleUri: vi.fn(),
    encodeRemoteModuleUri: vi.fn(),
    LiveShareIntegration: class LiveShareIntegration {},
    XLIDE_LIVESHARE_AUTHORITY: 'liveshare',
}));

import { decodeModuleUri, XlideFileSystemProvider } from '../src/xlideFileSystem';

/** Minimal stand-in — decodeModuleUri only reads uri.path */
function fakeUri(uriPath: string): VscodeType.Uri {
    return { path: uriPath, toString: () => uriPath } as VscodeType.Uri;
}

function moduleUriPath(workbookPath: string, moduleName = 'Module1'): string {
    const forward = workbookPath.replace(/\\/g, '/');
    const base = forward.startsWith('/') ? forward : `/${forward}`;
    return `${base}/${moduleName}.bas`;
}

describe('decodeModuleUri', () => {
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

    it('xlsmPath ends with the workbook filename', () => {
        const { xlsmPath } = decodeModuleUri(fakeUri('/home/user/mybook.xlsm/Module1.bas'));
        expect(xlsmPath.endsWith('mybook.xlsm')).toBe(true);
    });

    it('throws on a path with no recognised workbook extension', () => {
        expect(() => decodeModuleUri(fakeUri('/home/user/file.txt'))).toThrow();
    });

    it('throws on a path with missing module segment', () => {
        expect(() => decodeModuleUri(fakeUri('/home/user/workbook.xlsm/'))).toThrow();
    });
});

describe('XlideFileSystemProvider stats', () => {
    it('uses backing workbook mtime as the initial module mtime across provider restarts', () => {
        const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-stat-'));
        try {
            const workbook = path.join(dir, 'book.xlsm');
            fs.writeFileSync(workbook, 'fake workbook');
            const timestamp = new Date('2024-01-02T03:04:05.000Z');
            fs.utimesSync(workbook, timestamp, timestamp);
            const uri = fakeUri(moduleUriPath(workbook));

            const firstProvider = new XlideFileSystemProvider({ call: vi.fn() } as never);
            const secondProvider = new XlideFileSystemProvider({ call: vi.fn() } as never);

            const first = firstProvider.stat(uri);
            const second = secondProvider.stat(uri);
            expect(first.mtime).toBe(timestamp.getTime());
            expect(second.mtime).toBe(first.mtime);
        } finally {
            fs.rmSync(dir, { recursive: true, force: true });
        }
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
