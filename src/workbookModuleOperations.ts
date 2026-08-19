import * as vscode from 'vscode';
import type { WorkbookEngine } from './workbookEngine';
import type { XlsmExplorer } from './xlsmExplorer';
import type { VbaSymbolIndex } from './vbaSymbolIndex';
import {
    encodeModuleUri,
    notifySignatureDropped,
    type XlideFileSystemProvider,
} from './xlideFileSystem';
import { invalidateVbaMemberCompletionCache } from './vbaMemberCompletion';
import { runWriteWithExcelCoordination } from './excelWorkbookCoordinator';
import { noteModuleWrite } from './vbaRenameHistory';
import {
    discardPendingAgentReview,
    renamePendingAgentReview,
    trackModuleWriteForAgentReview,
} from './xlideAgentDiff';

/**
 * Shared workbook module mutations used by both the command handlers and the
 * agent (language-model) tools. Every write/rename/delete goes through one
 * code path: bridge call + signature-dropped notice + file-change event for
 * open editors + (for delete) closing stale tabs + project-state refresh.
 *
 * Audit records and user-facing messaging intentionally stay with the
 * callers - commands and agent tools present outcomes differently.
 */
export interface WorkbookModuleOperationDeps {
    bridge: WorkbookEngine;
    explorer: XlsmExplorer;
    fsProvider: XlideFileSystemProvider;
    vbaIndex: VbaSymbolIndex;
}

export interface WorkbookModuleMutationResult {
    ok?: boolean;
    signatureDropped?: boolean;
}

export interface WorkbookModuleOperationOptions {
    /**
     * Invalidate the per-workbook symbol/completion caches and refresh the
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

/** Drops cached per-workbook project state and refreshes the explorer tree. */
export function refreshWorkbookProjectState(
    deps: Pick<WorkbookModuleOperationDeps, 'explorer' | 'vbaIndex'>,
    filePath: string,
): void {
    deps.vbaIndex.invalidate(filePath);
    invalidateVbaMemberCompletionCache(filePath);
    deps.explorer.refresh();
}

export async function writeWorkbookModule(
    deps: WorkbookModuleOperationDeps,
    request: {
        filePath: string;
        moduleName: string;
        source: string;
        /** VBA module kind for the backend (e.g. 'standard', 'class'). */
        kind?: string;
    },
    options: WorkbookModuleOperationOptions = {},
): Promise<WorkbookModuleMutationResult> {
    const { filePath, moduleName, source, kind } = request;
    // Any other write makes the recorded rename's before-images stale: putting
    // them back would discard whatever this write is about to do. The undo path
    // takes the snapshot before it writes, so it is not tripped by its own
    // restores.
    noteModuleWrite(filePath, moduleName);
    const result = await runWriteWithExcelCoordination(filePath, () =>
        deps.bridge.call<WorkbookModuleMutationResult>('writeModule', {
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
        refreshWorkbookProjectState(deps, filePath);
    }
    return result;
}

/**
 * Writes a form's designer (control tree and textual properties) back into
 * the workbook from an imported `.frm`/`.frx` pair. Composes with
 * {@link writeWorkbookModule}, which owns the module's code.
 */
export async function writeWorkbookFormDesigner(
    deps: WorkbookModuleOperationDeps,
    request: {
        filePath: string;
        moduleName: string;
        frx: Buffer;
        frmDesignerBlock?: string;
    },
    options: WorkbookModuleOperationOptions = {},
): Promise<WorkbookModuleMutationResult> {
    const { filePath, moduleName, frx, frmDesignerBlock } = request;
    noteModuleWrite(filePath, moduleName);
    const result = await runWriteWithExcelCoordination(filePath, () =>
        deps.bridge.call<WorkbookModuleMutationResult>('writeFormDesigner', {
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
    deps.fsProvider.notifyFileChanged(encodeModuleUri(filePath, moduleName));
    if (options.refreshProjectState !== false) {
        refreshWorkbookProjectState(deps, filePath);
    }
    return result;
}

export async function renameWorkbookModule(
    deps: WorkbookModuleOperationDeps,
    request: { filePath: string; moduleName: string; newName: string },
    options: WorkbookModuleOperationOptions = {},
): Promise<WorkbookModuleMutationResult> {
    const { filePath, moduleName, newName } = request;
    const result = await runWriteWithExcelCoordination(filePath, () =>
        deps.bridge.call<WorkbookModuleMutationResult>('renameModule', {
            path: filePath,
            module: moduleName,
            newName,
        }),
    );
    notifySignatureDropped(filePath, Boolean(result.signatureDropped));
    // An unreviewed agent change follows the module to its new name.
    renamePendingAgentReview(filePath, moduleName, newName);
    // Tell open editors the old module is gone and refresh workbook stats
    deps.fsProvider.notifyFileChanged(encodeModuleUri(filePath, moduleName));
    if (options.refreshProjectState !== false) {
        refreshWorkbookProjectState(deps, filePath);
    }
    return result;
}

export async function deleteWorkbookModule(
    deps: WorkbookModuleOperationDeps,
    request: { filePath: string; moduleName: string },
    options: WorkbookModuleOperationOptions = {},
): Promise<WorkbookModuleMutationResult> {
    const { filePath, moduleName } = request;
    const result = await runWriteWithExcelCoordination(filePath, () =>
        deps.bridge.call<WorkbookModuleMutationResult>('deleteModule', {
            path: filePath,
            module: moduleName,
        }),
    );
    notifySignatureDropped(filePath, Boolean(result.signatureDropped));
    // A deleted module has nothing left to review.
    discardPendingAgentReview(filePath, moduleName);
    // Close any open editors for this module
    const uri = encodeModuleUri(filePath, moduleName);
    for (const tab of vscode.window.tabGroups.all.flatMap((g) => g.tabs)) {
        const input = tab.input;
        if (
            input instanceof vscode.TabInputText &&
            input.uri.toString() === uri.toString()
        ) {
            await vscode.window.tabGroups.close(tab);
        }
    }
    deps.fsProvider.notifyFileChanged(uri);
    if (options.refreshProjectState !== false) {
        refreshWorkbookProjectState(deps, filePath);
    }
    return result;
}
