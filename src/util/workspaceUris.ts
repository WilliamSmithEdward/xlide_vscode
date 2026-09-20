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
