import { describe, expect, it } from 'vitest';
import { moduleIdentityKey, projectIdentityKey, sameProjectPath } from '../src/projectIdentity';

// The key every surface compares projects by: the tree's active workbook, the
// engine's caches, the agent-review rows. Two strings naming one file have to
// produce one key, or the tree stops recognizing its own active project.
// Platform is a parameter here rather than ambient, so these run the same on
// any machine.

describe('win32', () => {
    it('folds case and separators, the way the filesystem does', () => {
        expect(projectIdentityKey('C:\\Work\\Book.xlsm', 'win32'))
            .toBe(projectIdentityKey('C:/WORK/book.xlsm', 'win32'));
        expect(sameProjectPath('C:\\Work\\Book.xlsm', 'c:\\work\\BOOK.XLSM', 'win32')).toBe(true);
    });

    it('resolves . and .. segments', () => {
        expect(projectIdentityKey('C:\\Work\\sub\\..\\Book.xlsm', 'win32'))
            .toBe(projectIdentityKey('C:\\Work\\Book.xlsm', 'win32'));
    });
});

describe('web', () => {
    // A browser hands the same workbook over in two forms: uri.fsPath renders
    // a virtual workspace's URI with backslashes, while decoding an
    // xlide-vba:// module URI keeps the URI's forward slashes. When these
    // keyed differently the tree could not match its active project, and
    // opening a module - or clicking its tab - collapsed the workbook.
    it('treats a backslash path and a slash path as one workbook', () => {
        expect(projectIdentityKey('\\Book.xlsm', 'web'))
            .toBe(projectIdentityKey('/Book.xlsm', 'web'));
        expect(sameProjectPath('\\Book.xlsm', '/Book.xlsm', 'web')).toBe(true);
    });

    it('does the same for a nested path', () => {
        expect(projectIdentityKey('\\repo\\src\\Book.xlsm', 'web'))
            .toBe(projectIdentityKey('/repo/src/Book.xlsm', 'web'));
    });

    it('keeps case, because a virtual filesystem may be case-sensitive', () => {
        expect(projectIdentityKey('/Book.xlsm', 'web'))
            .not.toBe(projectIdentityKey('/book.xlsm', 'web'));
    });

    it('still tells two different workbooks apart', () => {
        expect(sameProjectPath('/repo/A.xlsm', '/repo/B.xlsm', 'web')).toBe(false);
    });
});

describe('posix', () => {
    it('leaves a backslash alone: it is an ordinary character in a filename', () => {
        // Rewriting it would merge two files that genuinely differ on Linux.
        expect(projectIdentityKey('/repo/odd\\name.xlsm', 'linux'))
            .not.toBe(projectIdentityKey('/repo/odd/name.xlsm', 'linux'));
    });

    it('keeps case', () => {
        expect(sameProjectPath('/repo/Book.xlsm', '/repo/book.xlsm', 'linux')).toBe(false);
    });
});

describe('module identity', () => {
    it('folds case, since VBA module names are case-insensitive', () => {
        expect(moduleIdentityKey('Module1')).toBe(moduleIdentityKey('MODULE1'));
    });
});
