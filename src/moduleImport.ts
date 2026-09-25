// Applying an import plan: the one place module files from a folder go into
// a project. The Import Modules from Folder command applies the plan the
// user checked over in the preview, and the xlide_importModules agent tool
// applies the plan its confirmation described (issue #92). Both write
// through projectModuleOperations, so open editors, the tree and a pending
// agent review follow, and both leave the same audit entries.

import * as fs from 'fs';
import * as path from 'path';
import { splitFrmSource } from './vba/formDesigner';
import { withExportFolderLock } from './moduleExport';
import type { ModuleSyncPlan, ModuleSyncPlanItem } from './moduleSyncPlan';
import { updateProjectModuleSyncSettings } from './projectModuleSyncSettings';
import {
    deleteProjectModule,
    refreshProjectState,
    writeProjectFormDesigner,
    writeProjectModule,
    type ProjectModuleOperationDeps,
} from './projectModuleOperations';
import { errorMessage } from './util/errors';
import {
    formatChangeSummaryDetails,
    recordXlideWriteAuditEvent as recordWriteAudit,
} from './xlideWriteAudit';

export interface ImportApplyOptions {
    /** The command the audit entries name. */
    command: string;
    /** Every log line is tagged `[importModules]`, as the command's always were. */
    log: (line: string) => void;
    /**
     * As in writeProjectModule: the caller presents each write for review
     * itself, from the before-images this returns.
     */
    agentReviewHandled?: boolean;
}

/** A module the import wrote, with what it held before, for a review. */
export interface ImportedModule {
    moduleName: string;
    relativeName: string;
    /** The module's code before the import, as readModule gives it; empty for a module it created. */
    before: string;
    beforeExisted: boolean;
}

export interface ImportSkip {
    moduleName: string;
    relativeName: string;
    reason: string;
}

export interface ImportApplyResult {
    summary: string;
    written: ImportedModule[];
    removed: string[];
    skipped: ImportSkip[];
    failed: string[];
}

/** The plan's items whose ids are named, in the plan's order. */
export function selectedModuleSyncItems(
    plan: ModuleSyncPlan,
    selectedIds: readonly string[],
): ModuleSyncPlanItem[] {
    const selected = new Set(selectedIds);
    return plan.items.filter((item) => selected.has(item.id));
}

/**
 * Writes, creates and (in true-up) deletes the selected modules, refreshes
 * the project once, and records the folder and mode in the project's
 * settings. One module failing does not stop the others: each is its own
 * audit entry, and the summary counts them all.
 */
export async function applyImportModuleSyncPlan(
    deps: ProjectModuleOperationDeps,
    plan: ModuleSyncPlan,
    selectedIds: readonly string[],
    options: ImportApplyOptions,
): Promise<ImportApplyResult> {
    const { command, log } = options;
    const tag = 'importModules';
    const written: ImportedModule[] = [];
    const removed: string[] = [];
    const skipped: ImportSkip[] = [];
    const failed: string[] = [];

    for (const item of selectedModuleSyncItems(plan, selectedIds)) {
        if (item.status === 'unchanged') {
            skipped.push({ moduleName: item.moduleName, relativeName: item.relativeName, reason: 'unchanged' });
            continue;
        }
        if (item.status === 'will-remove') {
            try {
                log(`[${tag}] Deleting project module ${item.moduleName} during import true-up`);
                await deleteProjectModule(deps, {
                    filePath: plan.projectPath,
                    moduleName: item.moduleName,
                }, { refreshProjectState: false });
                removed.push(item.relativeName);
                recordWriteAudit({
                    command,
                    operation: 'delete-module',
                    outcome: 'succeeded',
                    projectPath: plan.projectPath,
                    moduleName: item.moduleName,
                    summary: 'Import true-up: 1 removed',
                });
            } catch (err) {
                failed.push(item.relativeName);
                recordWriteAudit({
                    command,
                    operation: 'delete-module',
                    outcome: 'failed',
                    projectPath: plan.projectPath,
                    moduleName: item.moduleName,
                    summary: 'Import true-up: 0 removed, 1 failed',
                    error: err,
                });
                log(`[${tag}] Error deleting ${item.moduleName}: ${errorMessage(err)}`);
            }
            continue;
        }
        if (item.status === 'skipping-import' || (item.unsupportedDirectCreation && !item.existsInProject)) {
            skipped.push({
                moduleName: item.moduleName,
                relativeName: item.relativeName,
                reason: `${item.moduleType} cannot be created directly`,
            });
            recordWriteAudit({
                command,
                operation: 'import-module',
                outcome: 'skipped',
                projectPath: plan.projectPath,
                moduleName: item.moduleName,
                sourcePath: item.sourcePath,
                summary: 'Import module: 0 changed, 1 skipped',
            });
            continue;
        }

        try {
            if (!item.sourcePath) {
                throw new Error(`Missing source path for ${item.moduleName}.`);
            }
            const sourcePath = item.sourcePath;
            // Read under the folder lock so a concurrent export cannot have a
            // half-written file in flight when we read it.
            const source = await withExportFolderLock(plan.folderPath, () =>
                fs.promises.readFile(sourcePath, 'utf8'));
            const before = options.agentReviewHandled
                ? await moduleBefore(deps, plan.projectPath, item.moduleName)
                : { before: '', beforeExisted: item.existsInProject };
            log(`[${tag}] Importing ${item.moduleName} from ${item.relativeName}`);
            await writeProjectModule(deps, {
                filePath: plan.projectPath,
                moduleName: item.moduleName,
                source,
                kind: item.moduleType,
            }, { refreshProjectState: false, agentReviewHandled: options.agentReviewHandled });
            // A .frm carries the form's designer in a sibling .frx; when the
            // pair is present and the form exists, the designer travels too.
            if (/\.frm$/i.test(item.relativeName) && item.existsInProject) {
                const frxPath = sourcePath.replace(/\.frm$/i, '.frx');
                const frx = await withExportFolderLock(plan.folderPath, () =>
                    fs.promises.readFile(frxPath).catch(() => undefined));
                const designerBlock = splitFrmSource(source)?.designerBlock;
                if (frx) {
                    log(`[${tag}] Importing designer for ${item.moduleName} from ${path.basename(frxPath)}`);
                    await writeProjectFormDesigner(deps, {
                        filePath: plan.projectPath,
                        moduleName: item.moduleName,
                        frx,
                        frmDesignerBlock: designerBlock,
                    }, { refreshProjectState: false });
                }
            }
            written.push({ moduleName: item.moduleName, relativeName: item.relativeName, ...before });
            recordWriteAudit({
                command,
                operation: 'import-module',
                outcome: 'succeeded',
                projectPath: plan.projectPath,
                moduleName: item.moduleName,
                sourcePath: item.sourcePath,
                summary: 'Import module: 1 changed',
            });
        } catch (err) {
            failed.push(item.relativeName);
            recordWriteAudit({
                command,
                operation: 'import-module',
                outcome: 'failed',
                projectPath: plan.projectPath,
                moduleName: item.moduleName,
                sourcePath: item.sourcePath,
                summary: 'Import module: 0 changed, 1 failed',
                error: err,
            });
            log(`[${tag}] Error importing ${item.moduleName}: ${errorMessage(err)}`);
        }
    }

    if (written.length > 0 || removed.length > 0) {
        refreshProjectState(deps, plan.projectPath);
    }
    try {
        await updateProjectModuleSyncSettings(plan.projectPath, {
            folderPath: plan.folderPath,
            exportMode: plan.exportMode,
            importMode: plan.importMode,
        });
    } catch (err) {
        failed.push('project settings');
        recordWriteAudit({
            command,
            operation: 'configure-module-sync',
            outcome: 'failed',
            projectPath: plan.projectPath,
            targetPath: plan.folderPath,
            summary: 'Sync settings: 0 changed, 1 failed',
            error: err,
        });
    }
    const lines = formatChangeSummaryDetails({
        operation: 'Import modules',
        changed: written.map((module) => module.relativeName),
        skipped: skipped.map((skip) => `${skip.relativeName} (${skip.reason})`),
        removed,
        failed,
    });
    for (const line of lines) {
        log(`[${tag}] ${line}`);
    }
    return { summary: lines[0], written, removed, skipped, failed };
}

/** What a module holds before the import writes it; a module the import creates holds nothing. */
async function moduleBefore(
    deps: Pick<ProjectModuleOperationDeps, 'bridge'>,
    projectPath: string,
    moduleName: string,
): Promise<{ before: string; beforeExisted: boolean }> {
    try {
        const read = await deps.bridge.call<{ source: string }>('readModule', { path: projectPath, module: moduleName });
        return { before: read.source, beforeExisted: true };
    } catch {
        return { before: '', beforeExisted: false };
    }
}
