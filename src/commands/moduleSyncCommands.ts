import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import {
    activeLocalVbaEditor,
    decodeModuleUri,
} from '../xlideFileSystem';
import {
    exportProjectModule,
    withExportFolderLock,
} from '../moduleExport';
import {
    applyImportModuleSyncPlan as applyImportPlan,
    selectedModuleSyncItems,
} from '../moduleImport';
import {
    buildExportModuleSyncPlan,
    buildImportModuleSyncPlan,
    type ModuleSyncPlan,
} from '../moduleSyncPlan';
import {
    openModuleSyncPreview,
    type ModuleSyncApplyResult,
    type ModuleSyncSettings,
    settingsFromPlan,
} from '../moduleSyncWebview';
import {
    effectiveProjectModuleSyncSettings,
    updateProjectModuleSyncSettings,
    type ProjectModuleSyncFolderSource,
    type ProjectModuleSyncModeSource,
} from '../projectModuleSyncSettings';
import {
    recordXlideWriteAuditEvent as recordWriteAudit,
    type XlideChangeSummary,
} from '../xlideWriteAudit';
import { registerXlideCommand } from '../xlideCommandRegistration';
import type { XlideNode } from '../projectExplorer';
import { errorMessage } from '../util/errors';
import { fileExists, isPathInside } from '../util/fs';
import {
    activeLocalProjectPath,
    logChangeSummary,
    resolveProjectPath,
    statusMessage,
    type CommandDeps,
    outputLogger,
} from './shared';

interface ResolvedModuleSyncSettings extends ModuleSyncSettings {
    folderPathSource: ProjectModuleSyncFolderSource;
    exportModeSource?: ProjectModuleSyncModeSource;
    importModeSource?: ProjectModuleSyncModeSource;
    settingsPath: string;
}

export function registerModuleSyncCommands(deps: CommandDeps): vscode.Disposable[] {
    const { context, bridge, out } = deps;

    const log = outputLogger(out);

    async function resolveModuleSyncFolder(
        filePath: string,
        direction: 'export' | 'import',
        options: { promptIfMissing?: boolean; openLabel?: string } = {},
    ): Promise<ResolvedModuleSyncSettings | undefined> {
        const existing = await effectiveProjectModuleSyncSettings(filePath);
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

    async function buildExportSyncPlanFromSettings(
        filePath: string,
        settings: ModuleSyncSettings,
    ): Promise<ModuleSyncPlan> {
        log(`[exportModules] Target folder: ${settings.folderPath}`);
        log(`[exportModules] Mode: ${settings.exportMode ?? 'exportAll'}`);
        return buildExportModuleSyncPlan(bridge, {
            projectPath: filePath,
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
            projectPath: filePath,
            importFolder: settings.folderPath,
            importMode: settings.importMode,
            folderPathSource: settings.folderPathSource,
            importModeSource: settings.importModeSource,
            settingsPath: settings.settingsPath,
        });
    }

    async function buildExportSyncPlanFromProjectSettings(
        filePath: string,
    ): Promise<ModuleSyncPlan | undefined> {
        const settings = await resolveModuleSyncFolder(filePath, 'export');
        return settings ? buildExportSyncPlanFromSettings(filePath, settings) : undefined;
    }

    async function buildImportSyncPlanFromProjectSettings(
        filePath: string,
    ): Promise<ModuleSyncPlan | undefined> {
        const settings = await resolveModuleSyncFolder(filePath, 'import');
        return settings ? buildImportSyncPlanFromSettings(filePath, settings) : undefined;
    }

    async function persistModuleSyncSettings(
        filePath: string,
        settings: ModuleSyncSettings,
    ): Promise<string> {
        const updated = await updateProjectModuleSyncSettings(filePath, {
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
            projectPath: filePath,
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
            vscode.window.showWarningMessage('XLIDE: Open a local VBA module to export the current module.');
            return;
        }

        if (editor.document.isDirty) {
            const saved = await editor.document.save();
            if (!saved) {
                vscode.window.showWarningMessage('XLIDE: Save the current module before exporting it.');
                return;
            }
        }

        const { projectPath, moduleName } = decodeModuleUri(editor.document.uri);
        const target = await resolveModuleSyncFolder(projectPath, 'export', { promptIfMissing: true, openLabel: 'Select export folder' });
        if (!target) {
            return;
        }

        log(`[exportCurrentModule] File: ${projectPath}`);
        log(`[exportCurrentModule] Module: ${moduleName}`);
        log(`[exportCurrentModule] Target folder: ${target.folderPath}`);
        log(`[exportCurrentModule] Mode: ${target.exportMode}`);

        const result = await exportProjectModule(bridge, {
            filePath: projectPath,
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
            projectPath: projectPath,
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

        log(`[exportModules] File: ${filePath}`);
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
                onReloadProjectSettings: () => buildExportSyncPlanFromProjectSettings(filePath),
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

        log(`[importModules] File: ${filePath}`);
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
                onReloadProjectSettings: () => buildImportSyncPlanFromProjectSettings(filePath),
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
                const result = await exportProjectModule(bridge, {
                    filePath: plan.projectPath,
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
            await persistModuleSyncSettings(plan.projectPath, settingsFromPlan(plan));
        } catch (err) {
            failed.push('project settings');
            recordWriteAudit({
                command: 'xlide.exportModulesToFolder',
                operation: 'configure-module-sync',
                outcome: 'failed',
                projectPath: plan.projectPath,
                targetPath: plan.folderPath,
                summary: 'Sync settings: 0 changed, 1 failed',
                error: err,
            });
            log(`[exportModules] Error updating project settings: ${errorMessage(err)}`);
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
            projectPath: plan.projectPath,
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
        const result = await applyImportPlan(deps, plan, selectedIds, {
            command: 'xlide.importModulesFromFolder',
            log,
        });
        return {
            summary: result.summary,
            changed: result.written.length,
            skipped: result.skipped.length,
            removed: result.removed.length,
            failed: result.failed.length,
        };
    }

    return [
        // Export all modules to a user-selected folder and persist folder in project config JSON
        registerXlideCommand('xlide.exportModulesToFolder', async (node: XlideNode) => {
            const filePath = resolveProjectPath(node);
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
                projectPath: resolveProjectPath(node),
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
                projectPath: await activeLocalProjectPath(),
                summary: 'Export current module: 0 changed, 1 failed',
                error: err,
            }),
        }),

        // Import selected module files from the configured (or user-chosen) export folder
        registerXlideCommand('xlide.importModulesFromFolder', async (node: XlideNode) => {
            const filePath = resolveProjectPath(node);
            if (!filePath) { return; }
            await showImportModulesDiffGui(filePath);
        }, { errorPrefix: 'Import failed', logTag: 'importModules', log }),
    ];
}
