import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import {
    activeLocalVbaEditor,
    decodeModuleUri,
} from '../xlideFileSystem';
import {
    exportWorkbookModule,
    withExportFolderLock,
} from '../moduleExport';
import {
    buildExportModuleSyncPlan,
    buildImportModuleSyncPlan,
    type ModuleSyncPlan,
    type ModuleSyncPlanItem,
} from '../moduleSyncPlan';
import {
    openModuleSyncPreview,
    type ModuleSyncApplyResult,
    type ModuleSyncSettings,
} from '../moduleSyncWebview';
import {
    effectiveWorkbookModuleSyncSettings,
    updateWorkbookModuleSyncSettings,
    type WorkbookModuleSyncFolderSource,
    type WorkbookModuleSyncModeSource,
} from '../workbookModuleSyncSettings';
import {
    recordXlideWriteAuditEvent as recordWriteAudit,
    type XlideChangeSummary,
} from '../xlideWriteAudit';
import {
    deleteWorkbookModule,
    refreshWorkbookProjectState,
    writeWorkbookModule,
} from '../workbookModuleOperations';
import { registerXlideCommand } from '../xlideCommandRegistration';
import type { XlideNode } from '../xlsmExplorer';
import { errorMessage } from '../util/errors';
import { fileExists, isPathInside } from '../util/fs';
import {
    activeLocalWorkbookPath,
    logChangeSummary,
    resolveWorkbookPath,
    statusMessage,
    type CommandDeps,
} from './shared';

interface ResolvedModuleSyncSettings extends ModuleSyncSettings {
    folderPathSource: WorkbookModuleSyncFolderSource;
    exportModeSource?: WorkbookModuleSyncModeSource;
    importModeSource?: WorkbookModuleSyncModeSource;
    settingsPath: string;
}

export function registerModuleSyncCommands(deps: CommandDeps): vscode.Disposable[] {
    const { context, bridge, out } = deps;

    function log(msg: string): void {
        out.appendLine(msg);
    }

    async function resolveModuleSyncFolder(
        filePath: string,
        direction: 'export' | 'import',
        options: { promptIfMissing?: boolean; openLabel?: string } = {},
    ): Promise<ResolvedModuleSyncSettings | undefined> {
        const existing = await effectiveWorkbookModuleSyncSettings(filePath);
        const modeFields = direction === 'export'
            ? { exportMode: existing.exportMode, exportModeSource: existing.exportModeSource }
            : { importMode: existing.importMode, importModeSource: existing.importModeSource };
        if (existing.folderPath) {
            return {
                folderPath: existing.folderPath,
                folderPathSource: existing.folderPathSource,
                ...modeFields,
                settingsPath: existing.settingsPath,
            };
        }
        if (!options.promptIfMissing) {
            return undefined;
        }

        const selected = await vscode.window.showOpenDialog({
            canSelectFiles: false,
            canSelectFolders: true,
            canSelectMany: false,
            openLabel: options.openLabel ?? 'Select folder to import from',
            defaultUri: vscode.Uri.file(path.dirname(filePath)),
        });
        return selected?.[0]?.fsPath ? {
            folderPath: selected[0].fsPath,
            folderPathSource: 'session',
            ...modeFields,
            settingsPath: existing.settingsPath,
        } : undefined;
    }

    async function chooseModuleSyncFolder(
        filePath: string,
        currentFolder: string | undefined,
        openLabel: string,
    ): Promise<string | undefined> {
        const selected = await vscode.window.showOpenDialog({
            canSelectFiles: false,
            canSelectFolders: true,
            canSelectMany: false,
            openLabel,
            defaultUri: currentFolder
                ? vscode.Uri.file(currentFolder)
                : vscode.Uri.file(path.dirname(filePath)),
        });
        return selected?.[0]?.fsPath;
    }

    function syncSettingsFromPlan(plan: ModuleSyncPlan): ModuleSyncSettings {
        return {
            folderPath: plan.folderPath,
            folderPathSource: plan.folderPathSource,
            exportMode: plan.exportMode,
            exportModeSource: plan.exportModeSource,
            importMode: plan.importMode,
            importModeSource: plan.importModeSource,
            settingsPath: plan.settingsPath,
        };
    }

    async function buildExportSyncPlanFromSettings(
        filePath: string,
        settings: ModuleSyncSettings,
    ): Promise<ModuleSyncPlan> {
        log(`[exportModules] Target folder: ${settings.folderPath}`);
        log(`[exportModules] Mode: ${settings.exportMode ?? 'exportAll'}`);
        return buildExportModuleSyncPlan(bridge, {
            workbookPath: filePath,
            exportFolder: settings.folderPath,
            exportMode: settings.exportMode,
            folderPathSource: settings.folderPathSource,
            exportModeSource: settings.exportModeSource,
            settingsPath: settings.settingsPath,
        });
    }

    async function buildImportSyncPlanFromSettings(
        filePath: string,
        settings: ModuleSyncSettings,
    ): Promise<ModuleSyncPlan> {
        log(`[importModules] Source folder: ${settings.folderPath}`);
        log(`[importModules] Mode: ${settings.importMode ?? 'updateOnly'}`);
        return buildImportModuleSyncPlan(bridge, {
            workbookPath: filePath,
            importFolder: settings.folderPath,
            importMode: settings.importMode,
            folderPathSource: settings.folderPathSource,
            importModeSource: settings.importModeSource,
            settingsPath: settings.settingsPath,
        });
    }

    async function buildExportSyncPlanFromWorkbookSettings(
        filePath: string,
    ): Promise<ModuleSyncPlan | undefined> {
        const settings = await resolveModuleSyncFolder(filePath, 'export');
        return settings ? buildExportSyncPlanFromSettings(filePath, settings) : undefined;
    }

    async function buildImportSyncPlanFromWorkbookSettings(
        filePath: string,
    ): Promise<ModuleSyncPlan | undefined> {
        const settings = await resolveModuleSyncFolder(filePath, 'import');
        return settings ? buildImportSyncPlanFromSettings(filePath, settings) : undefined;
    }

    async function persistModuleSyncSettings(
        filePath: string,
        settings: ModuleSyncSettings,
    ): Promise<string> {
        const updated = await updateWorkbookModuleSyncSettings(filePath, {
            folderPath: settings.folderPath,
            exportMode: settings.exportMode,
            importMode: settings.importMode,
        });
        return updated.settingsPath;
    }

    async function saveModuleSyncSettings(
        filePath: string,
        command: string,
        settings: ModuleSyncSettings,
    ): Promise<ModuleSyncApplyResult> {
        const configPath = await persistModuleSyncSettings(filePath, settings);
        const summary = 'Sync settings: 1 changed';
        log(`[moduleSyncSettings] Config updated: ${configPath}`);
        recordWriteAudit({
            command,
            operation: 'configure-module-sync',
            outcome: 'succeeded',
            workbookPath: filePath,
            targetPath: settings.folderPath,
            summary,
        });
        return {
            summary,
            changed: 1,
            skipped: 0,
            failed: 0,
        };
    }

    async function exportActiveModule(): Promise<void> {
        const editor = activeLocalVbaEditor();
        if (!editor) {
            vscode.window.showWarningMessage('XLIDE: Open a local workbook VBA module to export the current module.');
            return;
        }

        if (editor.document.isDirty) {
            const saved = await editor.document.save();
            if (!saved) {
                vscode.window.showWarningMessage('XLIDE: Save the current module before exporting it.');
                return;
            }
        }

        const { xlsmPath, moduleName } = decodeModuleUri(editor.document.uri);
        const target = await resolveModuleSyncFolder(xlsmPath, 'export', { promptIfMissing: true, openLabel: 'Select export folder' });
        if (!target) {
            return;
        }

        log(`[exportCurrentModule] Workbook: ${xlsmPath}`);
        log(`[exportCurrentModule] Module: ${moduleName}`);
        log(`[exportCurrentModule] Target folder: ${target.folderPath}`);
        log(`[exportCurrentModule] Mode: ${target.exportMode}`);

        const result = await exportWorkbookModule(bridge, {
            filePath: xlsmPath,
            moduleName,
            exportFolder: target.folderPath,
            exportMode: target.exportMode,
        });
        const changeSummary: XlideChangeSummary = {
            operation: 'Export current module',
            changed: result.writtenFiles,
        };
        const summaryText = logChangeSummary(log, 'exportCurrentModule', changeSummary);
        recordWriteAudit({
            command: 'xlide.exportCurrentModuleToFolder',
            operation: 'export-current-module',
            outcome: 'succeeded',
            workbookPath: xlsmPath,
            moduleName,
            targetPath: target.folderPath,
            summary: summaryText,
        });

        log(`[exportCurrentModule] Config updated: ${result.configPath}`);
        statusMessage(`XLIDE: ${summaryText} [mode=${result.exportMode}]`);
    }

    async function showExportModulesDiffGui(filePath: string): Promise<void> {
        const target = await resolveModuleSyncFolder(filePath, 'export', { promptIfMissing: true, openLabel: 'Select export folder' });
        if (!target) {
            return;
        }

        log(`[exportModules] Workbook: ${filePath}`);
        log(`[exportModules] Target folder: ${target.folderPath}`);
        log(`[exportModules] Mode: ${target.exportMode}`);

        const plan = await buildExportSyncPlanFromSettings(filePath, {
            folderPath: target.folderPath,
            folderPathSource: target.folderPathSource,
            exportMode: target.exportMode,
            exportModeSource: target.exportModeSource,
            settingsPath: target.settingsPath,
        });
        const result = await openModuleSyncPreview(
            context,
            plan,
            (currentPlan, selectedIds) => applyExportModuleSyncPlan(currentPlan, selectedIds),
            {
                onChooseFolder: async (settings) => {
                    const folderPath = await chooseModuleSyncFolder(filePath, settings.folderPath, 'Select export folder');
                    if (!folderPath) {
                        return undefined;
                    }
                    return buildExportSyncPlanFromSettings(filePath, { ...settings, folderPath, folderPathSource: 'session' });
                },
                onRefresh: (settings) => buildExportSyncPlanFromSettings(filePath, settings),
                onReloadWorkbookSettings: () => buildExportSyncPlanFromWorkbookSettings(filePath),
                onSaveSettings: (settings) => saveModuleSyncSettings(filePath, 'xlide.exportModulesToFolder', settings),
            },
        );
        if (!result) {
            return;
        }
        const message = `XLIDE: ${result.summary}`;
        if (result.failed > 0) {
            vscode.window.showWarningMessage(message);
        } else {
            statusMessage(message);
        }
    }

    async function showImportModulesDiffGui(filePath: string): Promise<void> {
        const target = await resolveModuleSyncFolder(filePath, 'import', { promptIfMissing: true });
        if (!target) {
            return;
        }

        log(`[importModules] Workbook: ${filePath}`);
        log(`[importModules] Source folder: ${target.folderPath}`);
        log(`[importModules] Mode: ${target.importMode}`);

        const plan = await buildImportSyncPlanFromSettings(filePath, {
            folderPath: target.folderPath,
            folderPathSource: target.folderPathSource,
            importMode: target.importMode,
            importModeSource: target.importModeSource,
            settingsPath: target.settingsPath,
        });
        const result = await openModuleSyncPreview(
            context,
            plan,
            (currentPlan, selectedIds) => applyImportModuleSyncPlan(currentPlan, selectedIds),
            {
                onChooseFolder: async (settings) => {
                    const folderPath = await chooseModuleSyncFolder(filePath, settings.folderPath, 'Select folder to import from');
                    if (!folderPath) {
                        return undefined;
                    }
                    return buildImportSyncPlanFromSettings(filePath, { ...settings, folderPath, folderPathSource: 'session' });
                },
                onRefresh: (settings) => buildImportSyncPlanFromSettings(filePath, settings),
                onReloadWorkbookSettings: () => buildImportSyncPlanFromWorkbookSettings(filePath),
                onSaveSettings: (settings) => saveModuleSyncSettings(filePath, 'xlide.importModulesFromFolder', settings),
            },
        );
        if (!result) {
            return;
        }
        if (result.failed > 0) {
            vscode.window.showWarningMessage(`XLIDE: ${result.summary}. Copy redacted diagnostics if you need to troubleshoot.`);
        } else {
            statusMessage(`XLIDE: ${result.summary} into ${path.basename(filePath)}`);
        }
    }

    async function applyExportModuleSyncPlan(
        plan: ModuleSyncPlan,
        selectedIds: readonly string[],
    ): Promise<ModuleSyncApplyResult> {
        const selected = selectedModuleSyncItems(plan, selectedIds);
        const changed: string[] = [];
        const skipped: string[] = [];
        const removed: string[] = [];
        const failed: string[] = [];

        for (const item of selected) {
            if (item.status === 'unchanged') {
                skipped.push(`${item.relativeName} (unchanged)`);
                continue;
            }
            if (item.status === 'will-remove') {
                try {
                    if (!item.targetPath || !isPathInside(plan.folderPath, item.targetPath)) {
                        throw new Error(`Refusing to remove a file outside the export folder: ${item.relativeName}`);
                    }
                    const targetPath = item.targetPath;
                    // Atomic check-then-delete under the export-folder lock, so a
                    // concurrent export cannot write this folder between the two.
                    const didRemove = await withExportFolderLock(plan.folderPath, async () => {
                        if (await fileExists(targetPath)) {
                            await fs.promises.unlink(targetPath);
                            return true;
                        }
                        return false;
                    });
                    if (didRemove) {
                        removed.push(item.relativeName);
                    } else {
                        skipped.push(`${item.relativeName} (already missing)`);
                    }
                } catch (err) {
                    failed.push(item.relativeName);
                    log(`[exportModules] Error removing ${item.relativeName}: ${errorMessage(err)}`);
                }
                continue;
            }

            try {
                const result = await exportWorkbookModule(bridge, {
                    filePath: plan.workbookPath,
                    moduleName: item.moduleName,
                    exportFolder: plan.folderPath,
                    exportMode: plan.exportMode,
                });
                changed.push(...result.writtenFiles);
            } catch (err) {
                failed.push(item.relativeName);
                log(`[exportModules] Error exporting ${item.moduleName}: ${errorMessage(err)}`);
            }
        }

        try {
            await persistModuleSyncSettings(plan.workbookPath, syncSettingsFromPlan(plan));
        } catch (err) {
            failed.push('workbook settings');
            recordWriteAudit({
                command: 'xlide.exportModulesToFolder',
                operation: 'configure-module-sync',
                outcome: 'failed',
                workbookPath: plan.workbookPath,
                targetPath: plan.folderPath,
                summary: 'Sync settings: 0 changed, 1 failed',
                error: err,
            });
            log(`[exportModules] Error updating workbook settings: ${errorMessage(err)}`);
        }
        const summaryText = logChangeSummary(log, 'exportModules', {
            operation: 'Export modules',
            changed,
            skipped,
            removed,
            failed,
        });
        recordWriteAudit({
            command: 'xlide.exportModulesToFolder',
            operation: 'export-modules',
            outcome: failed.length > 0 ? 'failed' : changed.length > 0 || removed.length > 0 ? 'succeeded' : 'skipped',
            workbookPath: plan.workbookPath,
            targetPath: plan.folderPath,
            summary: summaryText,
        });
        return {
            summary: summaryText,
            changed: changed.length,
            skipped: skipped.length,
            removed: removed.length,
            failed: failed.length,
        };
    }

    async function applyImportModuleSyncPlan(
        plan: ModuleSyncPlan,
        selectedIds: readonly string[],
    ): Promise<ModuleSyncApplyResult> {
        const selected = selectedModuleSyncItems(plan, selectedIds);
        const changed: string[] = [];
        const skipped: string[] = [];
        const removed: string[] = [];
        const failed: string[] = [];

        for (const item of selected) {
            if (item.status === 'unchanged') {
                skipped.push(`${item.relativeName} (unchanged)`);
                continue;
            }
            if (item.status === 'will-remove') {
                try {
                    log(`[importModules] Deleting workbook module ${item.moduleName} during import true-up`);
                    await deleteWorkbookModule(deps, {
                        filePath: plan.workbookPath,
                        moduleName: item.moduleName,
                    }, { refreshProjectState: false });
                    removed.push(item.relativeName);
                    recordWriteAudit({
                        command: 'xlide.importModulesFromFolder',
                        operation: 'delete-module',
                        outcome: 'succeeded',
                        workbookPath: plan.workbookPath,
                        moduleName: item.moduleName,
                        summary: 'Import true-up: 1 removed',
                    });
                } catch (err) {
                    failed.push(item.relativeName);
                    recordWriteAudit({
                        command: 'xlide.importModulesFromFolder',
                        operation: 'delete-module',
                        outcome: 'failed',
                        workbookPath: plan.workbookPath,
                        moduleName: item.moduleName,
                        summary: 'Import true-up: 0 removed, 1 failed',
                        error: err,
                    });
                    log(`[importModules] Error deleting ${item.moduleName}: ${errorMessage(err)}`);
                }
                continue;
            }
            if (item.status === 'skipping-import' || (item.unsupportedDirectCreation && !item.existsInWorkbook)) {
                skipped.push(`${item.relativeName} (${item.moduleType} cannot be created directly)`);
                recordWriteAudit({
                    command: 'xlide.importModulesFromFolder',
                    operation: 'import-module',
                    outcome: 'skipped',
                    workbookPath: plan.workbookPath,
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
                log(`[importModules] Importing ${item.moduleName} from ${item.relativeName}`);
                await writeWorkbookModule(deps, {
                    filePath: plan.workbookPath,
                    moduleName: item.moduleName,
                    source,
                    kind: item.moduleType,
                }, { refreshProjectState: false });
                changed.push(item.relativeName);
                recordWriteAudit({
                    command: 'xlide.importModulesFromFolder',
                    operation: 'import-module',
                    outcome: 'succeeded',
                    workbookPath: plan.workbookPath,
                    moduleName: item.moduleName,
                    sourcePath: item.sourcePath,
                    summary: 'Import module: 1 changed',
                });
            } catch (err) {
                failed.push(item.relativeName);
                recordWriteAudit({
                    command: 'xlide.importModulesFromFolder',
                    operation: 'import-module',
                    outcome: 'failed',
                    workbookPath: plan.workbookPath,
                    moduleName: item.moduleName,
                    sourcePath: item.sourcePath,
                    summary: 'Import module: 0 changed, 1 failed',
                    error: err,
                });
                log(`[importModules] Error importing ${item.moduleName}: ${errorMessage(err)}`);
            }
        }

        if (changed.length > 0 || removed.length > 0) {
            refreshWorkbookProjectState(deps, plan.workbookPath);
        }
        try {
            await persistModuleSyncSettings(plan.workbookPath, syncSettingsFromPlan(plan));
        } catch (err) {
            failed.push('workbook settings');
            recordWriteAudit({
                command: 'xlide.importModulesFromFolder',
                operation: 'configure-module-sync',
                outcome: 'failed',
                workbookPath: plan.workbookPath,
                targetPath: plan.folderPath,
                summary: 'Sync settings: 0 changed, 1 failed',
                error: err,
            });
        }
        const summaryText = logChangeSummary(log, 'importModules', {
            operation: 'Import modules',
            changed,
            skipped,
            removed,
            failed,
        });
        return {
            summary: summaryText,
            changed: changed.length,
            skipped: skipped.length,
            removed: removed.length,
            failed: failed.length,
        };
    }

    function selectedModuleSyncItems(
        plan: ModuleSyncPlan,
        selectedIds: readonly string[],
    ): ModuleSyncPlanItem[] {
        const selected = new Set(selectedIds);
        return plan.items.filter((item) => selected.has(item.id));
    }

    return [
        // Export all modules to a user-selected folder and persist folder in workbook config JSON
        registerXlideCommand('xlide.exportModulesToFolder', async (node: XlideNode) => {
            const filePath = resolveWorkbookPath(node);
            if (!filePath) { return; }
            await showExportModulesDiffGui(filePath);
        }, {
            errorPrefix: 'Failed to export modules',
            logTag: 'exportModules',
            log,
            onError: (err, node) => recordWriteAudit({
                command: 'xlide.exportModulesToFolder',
                operation: 'export-modules',
                outcome: 'failed',
                workbookPath: resolveWorkbookPath(node),
                summary: 'Export modules: 0 changed, 1 failed',
                error: err,
            }),
        }),

        // Save and export just the active VBA module to the configured module folder
        registerXlideCommand('xlide.exportCurrentModuleToFolder', () => exportActiveModule(), {
            errorPrefix: 'Failed to export current module',
            logTag: 'exportCurrentModule',
            log,
            onError: async (err) => recordWriteAudit({
                command: 'xlide.exportCurrentModuleToFolder',
                operation: 'export-current-module',
                outcome: 'failed',
                workbookPath: await activeLocalWorkbookPath(),
                summary: 'Export current module: 0 changed, 1 failed',
                error: err,
            }),
        }),

        // Import selected module files from the configured (or user-chosen) export folder
        registerXlideCommand('xlide.importModulesFromFolder', async (node: XlideNode) => {
            const filePath = resolveWorkbookPath(node);
            if (!filePath) { return; }
            await showImportModulesDiffGui(filePath);
        }, { errorPrefix: 'Import failed', logTag: 'importModules', log }),
    ];
}
