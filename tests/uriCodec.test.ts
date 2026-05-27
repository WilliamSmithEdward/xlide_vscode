import { vi, describe, it, expect } from 'vitest';
import type * as VscodeType from 'vscode';

vi.mock('vscode', () => ({}));
vi.mock('../src/pythonBridge', () => ({ PythonBridge: class PythonBridge {} }));
vi.mock('../src/liveShare', () => ({
    decodeRemoteModuleUri: vi.fn(),
    encodeRemoteModuleUri: vi.fn(),
    LiveShareIntegration: class LiveShareIntegration {},
    XLIDE_LIVESHARE_AUTHORITY: 'liveshare',
}));

import { decodeModuleUri } from '../src/xlideFileSystem';

/** Minimal stand-in — decodeModuleUri only reads uri.path */
function fakeUri(uriPath: string): VscodeType.Uri {
    return { path: uriPath } as VscodeType.Uri;
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
