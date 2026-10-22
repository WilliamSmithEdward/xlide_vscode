import { beforeEach, describe, expect, it, vi } from 'vitest';

const workspace = vi.hoisted(() => ({
    workspaceFolders: undefined as Array<{ uri: { fsPath: string; path: string } }> | undefined,
}));

vi.mock('vscode', () => ({ workspace }));

import { isInWorkspaceFolder } from '../src/util/workspaceUris';

/** A folder as VS Code reports it on Windows: lower-case drive in both forms. */
function windowsFolder(fsPath: string) {
    const lowered = fsPath.replace(/^[A-Z]:/, (drive) => drive.toLowerCase());
    return { uri: { fsPath: lowered, path: `/${lowered.replace(/\\/g, '/')}` } };
}

function posixFolder(folderPath: string) {
    return { uri: { fsPath: folderPath, path: folderPath } };
}

describe('isInWorkspaceFolder', () => {
    beforeEach(() => {
        workspace.workspaceFolders = undefined;
    });

    it('is false for every path when no folder is open', () => {
        expect(isInWorkspaceFolder('C:\\work\\Book.xlsm')).toBe(false);
        expect(isInWorkspaceFolder('/work/Book.xlsm')).toBe(false);
    });

    it('takes a file under a Windows folder whatever the drive letter case', () => {
        workspace.workspaceFolders = [windowsFolder('C:\\work')];
        expect(isInWorkspaceFolder('C:\\work\\Book.xlsm')).toBe(true);
        expect(isInWorkspaceFolder('c:\\work\\sub\\Book.xlsm')).toBe(true);
        expect(isInWorkspaceFolder('C:/Work/Book.xlsm')).toBe(true);
        expect(isInWorkspaceFolder('C:\\work')).toBe(true);
    });

    it('refuses a sibling folder that shares the prefix', () => {
        workspace.workspaceFolders = [windowsFolder('C:\\work')];
        expect(isInWorkspaceFolder('C:\\work2\\Book.xlsm')).toBe(false);
        expect(isInWorkspaceFolder('C:\\workbooks.xlsm')).toBe(false);
    });

    it('resolves .. before comparing', () => {
        workspace.workspaceFolders = [windowsFolder('C:\\work')];
        expect(isInWorkspaceFolder('C:\\work\\..\\secret\\Book.xlsm')).toBe(false);
        expect(isInWorkspaceFolder('C:\\work\\sub\\..\\Book.xlsm')).toBe(true);
        expect(isInWorkspaceFolder('C:\\work\\.\\Book.xlsm')).toBe(true);
    });

    it('refuses a relative path', () => {
        workspace.workspaceFolders = [windowsFolder('C:\\work')];
        expect(isInWorkspaceFolder('Book.xlsm')).toBe(false);
        expect(isInWorkspaceFolder('work\\Book.xlsm')).toBe(false);
    });

    it('keeps case on a POSIX folder', () => {
        workspace.workspaceFolders = [posixFolder('/home/me/work')];
        expect(isInWorkspaceFolder('/home/me/work/Book.xlsm')).toBe(true);
        expect(isInWorkspaceFolder('/home/me/Work/Book.xlsm')).toBe(false);
        expect(isInWorkspaceFolder('/home/me/work/../Book.xlsm')).toBe(false);
    });

    it('takes a file under any of several folders', () => {
        workspace.workspaceFolders = [posixFolder('/a'), windowsFolder('D:\\b')];
        expect(isInWorkspaceFolder('D:\\b\\Book.xlsm')).toBe(true);
        expect(isInWorkspaceFolder('/a/Book.xlsm')).toBe(true);
        expect(isInWorkspaceFolder('/c/Book.xlsm')).toBe(false);
    });

    it('takes a file under a UNC folder', () => {
        workspace.workspaceFolders = [{ uri: { fsPath: '\\\\server\\share\\work', path: '/work' } }];
        expect(isInWorkspaceFolder('\\\\server\\share\\work\\Book.xlsm')).toBe(true);
        expect(isInWorkspaceFolder('\\\\server\\share\\other\\Book.xlsm')).toBe(false);
    });

    it('takes a path under a virtual folder mounted at the root', () => {
        // github.dev: vscode-vfs://github/owner/repo, whose path is the repository.
        workspace.workspaceFolders = [{ uri: { fsPath: '/owner/repo', path: '/owner/repo' } }];
        expect(isInWorkspaceFolder('/owner/repo/Book.xlsm')).toBe(true);
        expect(isInWorkspaceFolder('/owner/other/Book.xlsm')).toBe(false);
    });
});
