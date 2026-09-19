import * as vscode from 'vscode';
import { createHash } from 'crypto';
import type { ProjectEngine } from './projectEngine';
import type { ProjectExplorer } from './projectExplorer';
import type { VbaSymbolIndex } from './vbaSymbolIndex';
import {
    encodeFormMarkupUri,
    encodeModuleUri,
    notifySignatureDropped,
    type XlideFileSystemProvider,
} from './xlideFileSystem';
import { invalidateVbaMemberCompletionCache } from './vbaMemberCompletion';
import { runWriteWithHostCoordination } from './officeWriteCoordinator';
import { onDidChangeProjectFile } from './projectFileChanges';
import { projectIdentityKey } from './projectIdentity';
import { noteModuleWrite } from './vbaRenameHistory';
import {
    discardPendingAgentReview,
    renamePendingAgentReview,
    trackModuleWriteForAgentReview,
} from './xlideAgentDiff';

/**
 * Shared project module mutations used by both the command handlers and the
 * agent (language-model) tools. Every write/rename/delete goes through one
 * code path: bridge call + signature-dropped notice + file-change event for
 * open editors + (for delete) closing stale tabs + project-state refresh.
 *
 * Audit records and user-facing messaging intentionally stay with the
 * callers - commands and agent tools present outcomes differently.
 */
export interface ProjectModuleOperationDeps {
    bridge: ProjectEngine;
    explorer: ProjectExplorer;
    fsProvider: XlideFileSystemProvider;
    vbaIndex: VbaSymbolIndex;
}

export interface ProjectModuleMutationResult {
    ok?: boolean;
    signatureDropped?: boolean;
    /**
     * The module a rename made, which is not always the name asked for: an
     * Access form's module keeps its `Form_` prefix.
     */
    moduleName?: string;
}

export interface ProjectModuleOperationOptions {
    /**
     * Invalidate the per-project symbol/completion caches and refresh the
     * explorer after the mutation (default). Batch callers (module sync)
     * pass false and refresh once after their loop.
     */
    refreshProjectState?: boolean;
    /**
     * The caller manages the agent review itself: the agent write tool (it
     * presents the review right after) and the review's own Revert (it
     * resolves the review right after). Every other write is tracked so a
     * pending review's after-image follows the module's live content.
     */
    agentReviewHandled?: boolean;
}

/** Drops cached per-project state and refreshes the explorer tree. */
export function refreshProjectState(
    deps: Pick<ProjectModuleOperationDeps, 'explorer' | 'vbaIndex'>,
    filePath: string,
): void {
    deps.vbaIndex.invalidate(filePath);
    invalidateVbaMemberCompletionCache(filePath);
    deps.explorer.refresh();
}

/**
 * Refreshes a project's cached state when its file changed outside XLIDE and
 * its VBA changed with it. Excel saving cells - AutoSave does that every few
 * seconds - and Access writing a database it merely has open change the file
 * without touching a module, and refreshing the tree and analyzing every open
 * module again for each was churn that could also fold rows the user had
 * opened. Open module documents follow on their own, in the file system
 * provider.
 */
export function refreshProjectStateOnOutsideChange(
    deps: Pick<ProjectModuleOperationDeps, 'bridge' | 'explorer' | 'vbaIndex'>,
): vscode.Disposable {
    // Each project's VBA as the last outside change found it. Any change XLIDE
    // makes clears it, so an outside change is never compared with VBA that
    // XLIDE has replaced since.
    const lastSeen = new Map<string, string>();
    const queues = new Map<string, Promise<void>>();
    const look = async (projectPath: string): Promise<void> => {
        let seen: string | undefined;
        try {
            const modules = await deps.bridge.call<unknown>('readModules', { path: projectPath });
            seen = createHash('sha256').update(JSON.stringify(modules)).digest('hex');
        } catch {
            // Unreadable for the moment: refresh, and compare from scratch next time.
        }
        const key = projectIdentityKey(projectPath);
        if (seen !== undefined && lastSeen.get(key) === seen) {
            return;
        }
        refreshProjectState(deps, projectPath);
        if (seen === undefined) {
            lastSeen.delete(key);
        } else {
            lastSeen.set(key, seen);
        }
    };
    return vscode.Disposable.from(
        onDidChangeProjectFile((projectPath) => {
            const key = projectIdentityKey(projectPath);
            const run = (queues.get(key) ?? Promise.resolve()).then(() => look(projectPath)).catch(() => undefined);
            queues.set(key, run);
        }),
        deps.vbaIndex.onDidChange(({ projectPath }) => {
            if (projectPath) {
                lastSeen.delete(projectIdentityKey(projectPath));
            } else {
                lastSeen.clear();
            }
        }),
    );
}

export async function writeProjectModule(
    deps: ProjectModuleOperationDeps,
    request: {
        filePath: string;
        moduleName: string;
        source: string;
        /** VBA module kind for the backend (e.g. 'standard', 'class'). */
        kind?: string;
    },
    options: ProjectModuleOperationOptions = {},
): Promise<ProjectModuleMutationResult> {
    const { filePath, moduleName, source, kind } = request;
    // Any other write makes the recorded rename's before-images stale: putting
    // them back would discard whatever this write is about to do. The undo path
    // takes the snapshot before it writes, so it is not tripped by its own
    // restores.
    noteModuleWrite(filePath, moduleName);
    const result = await runWriteWithHostCoordination(filePath, () =>
        deps.bridge.call<ProjectModuleMutationResult>('writeModule', {
            path: filePath,
            module: moduleName,
            source,
            ...(kind !== undefined ? { kind } : {}),
        }),
    );
    // Defensive: the backend signals failure by rejecting, but an explicit
    // ok:false must never be treated as success (silent no-op / data loss).
    if (result.ok === false) {
        throw new Error(`XLIDE: Writing module "${moduleName}" did not complete.`);
    }
    notifySignatureDropped(filePath, Boolean(result.signatureDropped));
    if (!options.agentReviewHandled) {
        trackModuleWriteForAgentReview(filePath, moduleName, source);
    }
    // Notify VS Code that the file changed so open editors reload
    deps.fsProvider.notifyFileChanged(encodeModuleUri(filePath, moduleName));
    if (options.refreshProjectState !== false) {
        refreshProjectState(deps, filePath);
    }
    return result;
}

/**
 * Writes a form's designer (control tree and textual properties) back into
 * the project from an imported `.frm`/`.frx` pair. Composes with
 * {@link writeProjectModule}, which owns the module's code.
 */
export async function writeProjectFormDesigner(
    deps: ProjectModuleOperationDeps,
    request: {
        filePath: string;
        moduleName: string;
        frx: Buffer;
        frmDesignerBlock?: string;
    },
    options: ProjectModuleOperationOptions = {},
): Promise<ProjectModuleMutationResult> {
    const { filePath, moduleName, frx, frmDesignerBlock } = request;
    noteModuleWrite(filePath, moduleName);
    const result = await runWriteWithHostCoordination(filePath, () =>
        deps.bridge.call<ProjectModuleMutationResult>('writeFormDesigner', {
            path: filePath,
            module: moduleName,
            frxBase64: frx.toString('base64'),
            ...(frmDesignerBlock !== undefined ? { frmDesignerBlock } : {}),
        }),
    );
    if (result.ok === false) {
        throw new Error(`XLIDE: Writing the designer of "${moduleName}" did not complete.`);
    }
    notifySignatureDropped(filePath, Boolean(result.signatureDropped));
    // The designer is what the form's markup document shows; its code is
    // written, and announced, by writeProjectModule.
    deps.fsProvider.notifyFileChanged(encodeFormMarkupUri(filePath, moduleName));
    if (options.refreshProjectState !== false) {
        refreshProjectState(deps, filePath);
    }
    return result;
}

/**
 * The addresses a module is shown at: its code, and for a form its markup,
 * which the form designer opens over as well.
 */
function moduleFaceUris(filePath: string, moduleName: string): vscode.Uri[] {
    return [encodeModuleUri(filePath, moduleName), encodeFormMarkupUri(filePath, moduleName)];
}

/** The tab's document address when it shows one of `uris` as text or in a custom editor. */
function tabFaceUri(tab: vscode.Tab, uris: readonly vscode.Uri[]): vscode.Uri | undefined {
    const input = tab.input;
    const uri = input instanceof vscode.TabInputText || input instanceof vscode.TabInputCustom ? input.uri : undefined;
    return uri && uris.some((candidate) => candidate.toString() === uri.toString()) ? uri : undefined;
}

/**
 * Moves the editors on a renamed module to its new name, each in its own group
 * and at its own selection: the code, and for a form its markup and designer.
 * An editor left on the old name showed a module that no longer exists. One
 * with unsaved edits stays: closing it would ask to save a module that is
 * gone, and its next save stops at VS Code's "file is newer" prompt instead.
 */
async function followRenamedModuleEditors(filePath: string, moduleName: string, newName: string): Promise<void> {
    const from = moduleFaceUris(filePath, moduleName);
    const to = moduleFaceUris(filePath, newName);
    const tabs = vscode.window.tabGroups.all.flatMap((group) => group.tabs).filter((tab) => !tab.isDirty);
    for (const tab of tabs) {
        const old = tabFaceUri(tab, from);
        if (!old) {
            continue;
        }
        const target = to[from.findIndex((uri) => uri.toString() === old.toString())];
        try {
            if (tab.input instanceof vscode.TabInputCustom) {
                await vscode.commands.executeCommand('vscode.openWith', target, tab.input.viewType, {
                    viewColumn: tab.group.viewColumn,
                    preview: tab.isPreview,
                    preserveFocus: true,
                });
            } else {
                const editor = vscode.window.visibleTextEditors.find((candidate) =>
                    candidate.document.uri.toString() === old.toString() && candidate.viewColumn === tab.group.viewColumn);
                // Markup is XML because Open Form Markup said so, not by its name.
                const languageId = vscode.workspace.textDocuments
                    .find((document) => document.uri.toString() === old.toString())?.languageId;
                const shown = await vscode.window.showTextDocument(target, {
                    viewColumn: tab.group.viewColumn,
                    preview: tab.isPreview,
                    preserveFocus: true,
                    selection: editor?.selection,
                });
                if (languageId && shown && shown.document.languageId !== languageId) {
                    await vscode.languages.setTextDocumentLanguage(shown.document, languageId);
                }
            }
            await vscode.window.tabGroups.close(tab, true);
        } catch {
            // Best-effort: an editor that cannot follow stays on the old name.
        }
    }
}

/** Tells the documents open on a module's faces that it changed. */
function notifyModuleFacesChanged(deps: Pick<ProjectModuleOperationDeps, 'fsProvider'>, filePath: string, moduleName: string): void {
    const [code, markup] = moduleFaceUris(filePath, moduleName);
    deps.fsProvider.notifyFileChanged(code);
    // Only a form has markup; a document of it is open only then.
    if (vscode.workspace.textDocuments.some((document) => document.uri.toString() === markup.toString())) {
        deps.fsProvider.notifyFileChanged(markup);
    }
}

export async function renameProjectModule(
    deps: ProjectModuleOperationDeps,
    request: { filePath: string; moduleName: string; newName: string },
    options: ProjectModuleOperationOptions = {},
): Promise<ProjectModuleMutationResult> {
    const { filePath, moduleName, newName } = request;
    const result = await runWriteWithHostCoordination(filePath, () =>
        deps.bridge.call<ProjectModuleMutationResult>('renameModule', {
            path: filePath,
            module: moduleName,
            newName,
        }),
    );
    notifySignatureDropped(filePath, Boolean(result.signatureDropped));
    // The module the engine made: `Customers` makes an Access form's module
    // `Form_Customers`.
    const renamedTo = result.moduleName ?? newName;
    // An unreviewed agent change follows the module to its new name, and so
    // do the editors on it.
    renamePendingAgentReview(filePath, moduleName, renamedTo);
    await followRenamedModuleEditors(filePath, moduleName, renamedTo);
    // Tell open editors the old module is gone and refresh project stats
    notifyModuleFacesChanged(deps, filePath, moduleName);
    if (options.refreshProjectState !== false) {
        refreshProjectState(deps, filePath);
    }
    return { ...result, moduleName: renamedTo };
}

export async function deleteProjectModule(
    deps: ProjectModuleOperationDeps,
    request: { filePath: string; moduleName: string },
    options: ProjectModuleOperationOptions = {},
): Promise<ProjectModuleMutationResult> {
    const { filePath, moduleName } = request;
    const result = await runWriteWithHostCoordination(filePath, () =>
        deps.bridge.call<ProjectModuleMutationResult>('deleteModule', {
            path: filePath,
            module: moduleName,
        }),
    );
    notifySignatureDropped(filePath, Boolean(result.signatureDropped));
    // A deleted module has nothing left to review.
    discardPendingAgentReview(filePath, moduleName);
    // Close any open editors for this module: its code, and for a form its
    // markup and designer, which otherwise went on showing a form that was gone.
    const faces = moduleFaceUris(filePath, moduleName);
    for (const tab of vscode.window.tabGroups.all.flatMap((g) => g.tabs)) {
        if (tabFaceUri(tab, faces)) {
            await vscode.window.tabGroups.close(tab);
        }
    }
    notifyModuleFacesChanged(deps, filePath, moduleName);
    if (options.refreshProjectState !== false) {
        refreshProjectState(deps, filePath);
    }
    return result;
}
