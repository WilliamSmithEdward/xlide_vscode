// Turning a project path back into the URI it came from.
//
// XLIDE identifies a project by a path string, and has since long before
// there was a browser build: the engine takes paths, the tree keys on them,
// the xlide-vba:// module URIs embed them. On a desktop that is lossless,
// because vscode.Uri.file() rebuilds the original exactly.
//
// A virtual workspace is not `file:`. github.dev serves
// vscode-vfs://github/owner/repo/Book.xlsm, and Uri.file() on that path
// produces a `file:` URI pointing at nothing - which is how "Failed to list
// modules: File not found" happens with the workbook plainly in the tree.
//
// Rather than thread vscode.Uri through every path in the extension, the URI
// is recovered from the workspace folder the path sits in. The folder knows
// its own scheme and authority, so a path under it reconstructs exactly.

import * as vscode from 'vscode';

/**
 * The URI a project path refers to.
 *
 * A path inside a workspace folder is rebuilt against that folder's own URI,
 * which carries the scheme and authority. Anything else falls back to a
 * `file:` URI, which is right on a desktop and the only guess available
 * elsewhere.
 */
export function workspaceUriFor(filePath: string): vscode.Uri {
    for (const folder of vscode.workspace.workspaceFolders ?? []) {
        const relative = relativeToFolder(folder.uri, filePath);
        if (relative !== undefined) {
            return folder.uri.with({
                path: relative.length > 0 ? `${trimTrailingSlash(folder.uri.path)}/${relative}` : folder.uri.path,
            });
        }
    }
    return vscode.Uri.file(filePath);
}

/**
 * The part of `filePath` below the folder, in POSIX form, or undefined when
 * it is not below the folder at all. Compared case-insensitively on the
 * separator only: the rest is compared as given, because a virtual
 * filesystem may well be case-sensitive.
 */
function relativeToFolder(folderUri: vscode.Uri, filePath: string): string | undefined {
    const normalized = filePath.replace(/\\/g, '/');
    for (const base of [folderUri.fsPath.replace(/\\/g, '/'), folderUri.path]) {
        const trimmed = trimTrailingSlash(base);
        if (normalized === trimmed) {
            return '';
        }
        // A folder at the root trims to '', which still prefixes every path
        // in it. test-web mounts exactly like that, and so does a repository
        // opened at its top level.
        const prefix = trimmed === '' ? '/' : `${trimmed}/`;
        if (normalized.startsWith(prefix)) {
            return normalized.slice(prefix.length);
        }
    }
    return undefined;
}

function trimTrailingSlash(value: string): string {
    return value.endsWith('/') ? value.slice(0, -1) : value;
}

/**
 * Whether `filePath` is in one of the workspace folders: what decides if an
 * agent may read it without asking.
 *
 * Unlike workspaceUriFor this is a boundary, so it fails closed. `..` and
 * `.` segments are resolved first, or `C:\work\..\secret.xlsm` would pass as
 * inside C:\work. A relative path is outside, since it resolves against
 * whatever directory the extension host runs in. A Windows drive path is
 * compared without case, because the file system ignores it and VS Code
 * lower-cases the drive letter of a folder's fsPath.
 */
export function isInWorkspaceFolder(filePath: string): boolean {
    const target = normalizedAbsolute(filePath);
    if (target === undefined) {
        return false;
    }
    for (const folder of vscode.workspace.workspaceFolders ?? []) {
        for (const base of [folder.uri.fsPath, folder.uri.path]) {
            const root = normalizedAbsolute(base);
            if (root !== undefined && containsPath(root, target)) {
                return true;
            }
        }
    }
    return false;
}

const WINDOWS_DRIVE = /^[a-zA-Z]:\//;

function containsPath(root: string, target: string): boolean {
    const fold = WINDOWS_DRIVE.test(root) && WINDOWS_DRIVE.test(target);
    const [r, t] = fold ? [root.toLowerCase(), target.toLowerCase()] : [root, target];
    if (t === r) {
        return true;
    }
    return t.startsWith(r.endsWith('/') ? r : `${r}/`);
}

/**
 * The path in forward-slash form with `.` and `..` resolved, or undefined
 * when it is not absolute. A `..` above the root stays at the root, as the
 * file system does.
 */
function normalizedAbsolute(value: string): string | undefined {
    const slashed = value.replace(/\\/g, '/');
    let prefix: string;
    if (slashed.startsWith('//')) {
        prefix = '//';
    } else if (WINDOWS_DRIVE.test(slashed)) {
        prefix = slashed.slice(0, 3);
    } else if (slashed.startsWith('/')) {
        prefix = '/';
    } else {
        return undefined;
    }
    const segments: string[] = [];
    for (const segment of slashed.slice(prefix.length).split('/')) {
        if (segment === '' || segment === '.') {
            continue;
        }
        if (segment === '..') {
            segments.pop();
        } else {
            segments.push(segment);
        }
    }
    return prefix + segments.join('/');
}
