// Notices a project file changing when XLIDE did not write it: the VBE saving
// the workbook, a git checkout or pull, the MCP server writing in its own
// process, a copy dropped over the file. XLIDE keeps what it read - open
// module documents, the tree's listing, the symbol index - and nothing
// followed such a change, so an open module went on showing the old code and
// analysis went on using it until the window was reloaded.
//
// Every XLIDE write records the file's stamp before and after it, so the stamp
// an XLIDE write leaves is XLIDE's own. Any other stamp fires
// onDidChangeProjectFile, found three ways:
//
//   - a watcher over every macro container in the workspace, which is what
//     the tree lists and so the case that has to work with nothing open;
//   - a watcher on one file, held while a module of it is open, which also
//     covers a project outside the workspace;
//   - the check the file system provider makes whenever VS Code asks it
//     about a module, before a read or a save.

import { hostPlatform } from './vba/hostPlatform';
import { workspaceUriFor } from './util/workspaceUris';
import * as path from 'path';
import * as vscode from 'vscode';
import { MACRO_CONTAINER_GLOB } from './macroContainerUi';
import { projectIdentityKey, sameProjectPath } from './projectIdentity';

/** Long enough for an application's save, which can touch the file more than once. */
const WATCH_SETTLE_MS = 200;

/** Modification time and size: what changes when anything rewrites the file. */
function stampOf(projectPath: string): string | undefined {
    try {
        const stat = hostPlatform().stat(projectPath);
        return `${stat.mtimeMs}:${stat.size}`;
    } catch {
        return undefined;
    }
}

/** The last stamp XLIDE accounted for, per project identity. */
const accounted = new Map<string, string>();
/** Checks waiting for a file to stop changing, per project identity. */
const settling = new Map<string, ReturnType<typeof setTimeout>>();
const changeEmitter = new vscode.EventEmitter<string>();

/** Fires with a project's path when its file changed and XLIDE did not write it. */
export const onDidChangeProjectFile = changeEmitter.event;

/**
 * Compares the file with the last stamp XLIDE accounted for, and fires when
 * they differ. The first look at a project only records it, since nothing was
 * read from it before. A file that cannot be statted fires nothing: an
 * application's save replaces the file through a rename, so it can be missing
 * for a moment, and the next look decides. True when the event fired.
 */
export function checkProjectFile(projectPath: string): boolean {
    const stamp = stampOf(projectPath);
    if (stamp === undefined) {
        return false;
    }
    const key = projectIdentityKey(projectPath);
    const known = accounted.get(key);
    accounted.set(key, stamp);
    if (known === undefined || known === stamp) {
        return false;
    }
    changeEmitter.fire(projectPath);
    return true;
}

/**
 * Checks the file once it has stopped changing, so one save is one check.
 * Shared by both watchers, so a file they both cover is checked once.
 */
export function scheduleProjectFileCheck(projectPath: string): void {
    const key = projectIdentityKey(projectPath);
    const pending = settling.get(key);
    if (pending) {
        clearTimeout(pending);
    }
    settling.set(key, setTimeout(() => {
        settling.delete(key);
        checkProjectFile(projectPath);
    }, WATCH_SETTLE_MS));
}

/**
 * Runs one XLIDE write to a project file and takes the stamp it leaves as
 * XLIDE's own. A file that had already changed under XLIDE before the write is
 * left for the next check to report: taking the stamp then would hide that
 * change from every module the write did not touch.
 */
export async function recordProjectWrite<T>(projectPath: string, write: () => Promise<T>): Promise<T> {
    const before = stampOf(projectPath);
    const result = await write();
    const after = stampOf(projectPath);
    const key = projectIdentityKey(projectPath);
    const known = accounted.get(key);
    if (after !== undefined && (known === undefined || known === before)) {
        accounted.set(key, after);
    }
    return result;
}

/**
 * Checks every macro container in the workspace whenever one changes on disk,
 * for the life of the returned disposable. Held once, for the session.
 *
 * This is the case the per-file watch below cannot cover, and the ordinary
 * one: a project the user is only looking at in the tree has no module
 * document open, so nothing armed a watch for it and nothing asked the file
 * system provider about it. A module renamed in the VBE and saved, or written
 * by the MCP server in its own process, reached nothing at all, and the tree
 * went on listing what the file held when it was last read.
 *
 * Both events lead to the same check because a save that replaces the file
 * through a rename can report either. Measured in a real VS Code (1.138,
 * Windows): writing a sibling temp file and renaming it over the target -
 * which is how XLIDE, Office and the MCP server all write - reports `change`,
 * so `onDidCreate` alone (which already refreshes the tree, for a file
 * arriving in the workspace) never saw a save.
 */
export function watchWorkspaceProjectFiles(): vscode.Disposable {
    const watcher = vscode.workspace.createFileSystemWatcher(MACRO_CONTAINER_GLOB);
    const check = (uri: vscode.Uri): void => scheduleProjectFileCheck(uri.fsPath);
    return vscode.Disposable.from(watcher.onDidChange(check), watcher.onDidCreate(check), watcher);
}

interface ProjectFileWatch {
    holders: number;
    subscriptions?: vscode.Disposable;
}

const watches = new Map<string, ProjectFileWatch>();

/**
 * Checks a project file whenever it changes on disk, for as long as the
 * returned disposable is held. Holders of one project share a watcher.
 *
 * This covers a project OUTSIDE the workspace, which the workspace-wide
 * watch cannot see: a module document can be opened for any file the tree
 * reached, and the settle in {@link scheduleProjectFileCheck} keeps the two
 * watches from checking the same save twice.
 */
export function watchProjectFile(projectPath: string): vscode.Disposable {
    const key = projectIdentityKey(projectPath);
    let held = watches.get(key);
    if (!held) {
        const watch: ProjectFileWatch = { holders: 0 };
        const settle = (uri: vscode.Uri): void => {
            if (sameProjectPath(uri.fsPath, projectPath)) {
                scheduleProjectFileCheck(projectPath);
            }
        };
        // The folder rather than the file: a file name can hold glob syntax
        // (`Book [1].xlsm`), and a pattern made from it would not match it.
        const watcher = vscode.workspace.createFileSystemWatcher(
            new vscode.RelativePattern(workspaceUriFor(path.dirname(projectPath)), '*'),
        );
        // A save that replaces the file through a rename reports a create.
        watch.subscriptions = vscode.Disposable.from(watcher.onDidChange(settle), watcher.onDidCreate(settle), watcher);
        watches.set(key, watch);
        held = watch;
    }
    const mine = held;
    mine.holders += 1;
    let released = false;
    return new vscode.Disposable(() => {
        if (released) {
            return;
        }
        released = true;
        mine.holders -= 1;
        if (mine.holders === 0) {
            mine.subscriptions?.dispose();
            if (watches.get(key) === mine) {
                watches.delete(key);
            }
        }
    });
}
