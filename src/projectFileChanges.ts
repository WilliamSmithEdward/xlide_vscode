// Notices a project file changing when XLIDE did not write it: the VBE saving
// the workbook, a git checkout or pull, a copy dropped over the file. XLIDE
// keeps what it read - open module documents, the symbol index - and nothing
// followed such a change, so an open module went on showing the old code and
// analysis went on using it until the window was reloaded.
//
// Every XLIDE write records the file's stamp before and after it, so the stamp
// an XLIDE write leaves is XLIDE's own. Any other stamp fires
// onDidChangeProjectFile, found by a watcher on the file while one of its
// modules is open, or by the check the file system provider makes whenever
// VS Code asks it about a module, before a read or a save.

import { hostPlatform } from './vba/hostPlatform';
import { workspaceUriFor } from './util/workspaceUris';
import * as path from 'path';
import * as vscode from 'vscode';
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

interface ProjectFileWatch {
    holders: number;
    timer?: ReturnType<typeof setTimeout>;
    subscriptions?: vscode.Disposable;
}

const watches = new Map<string, ProjectFileWatch>();

/**
 * Checks a project file whenever it changes on disk, for as long as the
 * returned disposable is held. Holders of one project share a watcher.
 */
export function watchProjectFile(projectPath: string): vscode.Disposable {
    const key = projectIdentityKey(projectPath);
    let held = watches.get(key);
    if (!held) {
        const watch: ProjectFileWatch = { holders: 0 };
        const settle = (uri: vscode.Uri): void => {
            if (!sameProjectPath(uri.fsPath, projectPath)) {
                return;
            }
            if (watch.timer) {
                clearTimeout(watch.timer);
            }
            watch.timer = setTimeout(() => {
                watch.timer = undefined;
                checkProjectFile(projectPath);
            }, WATCH_SETTLE_MS);
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
            if (mine.timer) {
                clearTimeout(mine.timer);
            }
            mine.subscriptions?.dispose();
            if (watches.get(key) === mine) {
                watches.delete(key);
            }
        }
    });
}
