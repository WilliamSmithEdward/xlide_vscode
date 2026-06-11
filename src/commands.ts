import * as vscode from 'vscode';
import * as path from 'path';
import * as cp from 'child_process';
import * as fs from 'fs';
import { PythonBridge } from './pythonBridge';
import { XlsmExplorer, XlideNode } from './xlsmExplorer';
import {
    XlideFileSystemProvider,
    encodeModuleUri,
    decodeModuleUri,
    XLIDE_SCHEME,
    XLIDE_VBA_LANGUAGE_ID,
    activeLocalVbaEditor,
} from './xlideFileSystem';
import { applyOpenDocumentSources } from './vbaOpenDocuments';
import { encodeRemoteModuleUri } from './liveShare';
import {
    exportWorkbookModule,
} from './moduleExport';
import {
    type ExportMode,
} from './workbookSettings';
import { xlideAttachToRunningExcelFromConfig } from './globalSettings';
import { validateVbaModuleName } from './vbaSourceScan';
import { VbaSymbolIndex } from './vbaSymbolIndex';
import { projectClassModuleDefinition } from './vbaNavigation';
import { buildVbaProjectIndexAsync } from './vbaProjectAnalysis';
import {
    projectClassReferenceEdit,
    renameProjectClassModule,
} from './vbaClassRename';
import {
    projectStandardModuleReferenceEdit,
    renameProjectStandardModule,
} from './vbaStandardModuleRename';
import { registerXlideCommand } from './xlideCommandRegistration';
import {
    recordXlideWriteAuditEvent as recordWriteAudit,
    type XlideChangeSummary,
} from './xlideWriteAudit';
import {
    buildExportModuleSyncPlan,
    buildImportModuleSyncPlan,
    type ImportMode,
    type ModuleSyncPlan,
    type ModuleSyncPlanItem,
} from './moduleSyncPlan';
import {
    openModuleSyncPreview,
    type ModuleSyncApplyResult,
    type ModuleSyncSettings,
} from './moduleSyncWebview';
import {
    effectiveWorkbookModuleSyncSettings,
    updateWorkbookModuleSyncSettings,
    type WorkbookModuleSyncFolderSource,
    type WorkbookModuleSyncModeSource,
} from './workbookModuleSyncSettings';
import {
    activeLocalWorkbookPath,
    logChangeSummary,
    procedureNameAtCursor,
    resolveWorkbookPath,
    type CommandDeps,
} from './commands/shared';
import { registerAnalysisCommands } from './commands/analysisCommands';
import { registerSupportBundleCommands } from './commands/supportBundleCommands';
import { registerVbaTestCommands } from './commands/vbaTestCommands';
import {
    deleteWorkbookModule,
    refreshWorkbookProjectState,
    writeWorkbookModule,
    type WorkbookModuleOperationDeps,
} from './workbookModuleOperations';
import { errorMessage } from './util/errors';
import { fileExists, isPathInside } from './util/fs';
import {
    ExcelMacroError,
    openWorkbookInExcel,
    runWorkbookMacroReadOnly,
} from './excelLauncher';

interface ResolvedModuleSyncSettings extends ModuleSyncSettings {
    folderPathSource: WorkbookModuleSyncFolderSource;
    exportModeSource?: WorkbookModuleSyncModeSource;
    importModeSource?: WorkbookModuleSyncModeSource;
    settingsPath: string;
}

export function registerCommands(
    context: vscode.ExtensionContext,
    bridge: PythonBridge,
    explorer: XlsmExplorer,
    fsProvider: XlideFileSystemProvider,
    out: vscode.OutputChannel,
    vbaIndex: VbaSymbolIndex,
): vscode.Disposable[] {
    const deps: CommandDeps = { context, bridge, explorer, fsProvider, out, vbaIndex };
    const moduleOps: WorkbookModuleOperationDeps = { bridge, explorer, fsProvider, vbaIndex };

    function log(msg: string): void {
        out.appendLine(msg);
    }

    function shouldAttachToRunningExcel(): boolean {
        return xlideAttachToRunningExcelFromConfig(vscode.workspace.getConfiguration('xlide')).value;
    }

    function showRunMacroFailure(err: unknown): void {
        if (err instanceof ExcelMacroError &&
            (err.code === 'REOPEN_BLOCKED' || err.code === 'REOPEN_FAILED')) {
            void vscode.window.showWarningMessage(`XLIDE: ${err.message}`);
            return;
        }
        void vscode.window.showErrorMessage(`XLIDE: Failed to run macro: ${errorMessage(err)}`);
    }

    // Windows COM-based Excel launch; the script lives in excelLauncher.ts.
    function runWindowsExcel(filePath: string, attachToRunning: boolean, readOnly: boolean): void {
        void openWorkbookInExcel(filePath, { attachToRunning, readOnly }, log).catch((err: Error) => {
            void vscode.window.showErrorMessage(`XLIDE: Open Workbook failed: ${err.message}`);
        });
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
        vscode.window.showInformationMessage(
            `XLIDE: ${summaryText} [mode=${result.exportMode}]`,
        );
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
            vscode.window.showInformationMessage(message);
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
            vscode.window.showInformationMessage(`XLIDE: ${result.summary} into ${path.basename(filePath)}`);
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
                    if (await fileExists(item.targetPath)) {
                        await fs.promises.unlink(item.targetPath);
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
                    await deleteWorkbookModule(moduleOps, {
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
                const source = await fs.promises.readFile(item.sourcePath, 'utf8');
                log(`[importModules] Importing ${item.moduleName} from ${item.relativeName}`);
                await writeWorkbookModule(moduleOps, {
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
            refreshWorkbookProjectState(moduleOps, plan.workbookPath);
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

    async function showClassModuleReferences(node: XlideNode): Promise<void> {
        if (!node.moduleName || !node.filePath || node.isRemote) { return; }
        const originUri = encodeModuleUri(node.filePath, node.moduleName);
        const originDoc = await vscode.workspace.openTextDocument(originUri);
        await vscode.languages.setTextDocumentLanguage(originDoc, XLIDE_VBA_LANGUAGE_ID);
        const editor = await vscode.window.showTextDocument(originDoc, { preview: false });
        const origin = new vscode.Position(0, 0);
        editor.selection = new vscode.Selection(origin, origin);
        await vscode.commands.executeCommand('references-view.findReferences', originUri, origin);
    }

    return [
        registerXlideCommand('xlide.refreshExplorer', () => {
            explorer.refresh();
        }),

        // Open a module (or navigate to a sub's line inside one)
        registerXlideCommand('xlide.openModule', async (node: XlideNode) => {
            if (!node?.moduleName) { return; }
            const uri = node.isRemote && node.remoteId
                ? encodeRemoteModuleUri(node.remoteId, node.moduleName)
                : encodeModuleUri(node.filePath, node.moduleName);

            const doc = await vscode.workspace.openTextDocument(uri);
            await vscode.languages.setTextDocumentLanguage(doc, XLIDE_VBA_LANGUAGE_ID);
            const editor = await vscode.window.showTextDocument(doc, { preview: false });

            // If a specific line was requested (sub navigation), move cursor there
            if (node.line !== undefined && node.line > 0) {
                const pos = new vscode.Position(node.line - 1, 0);
                editor.selection = new vscode.Selection(pos, pos);
                editor.revealRange(
                    new vscode.Range(pos, pos),
                    vscode.TextEditorRevealType.InCenterIfOutsideViewport,
                );
            }
        }),

        // Find all references to the procedure or class represented by a tree node
        registerXlideCommand('xlide.findReferences', async (node: XlideNode) => {
            if (!node?.moduleName) { return; }
            if (node.kind === 'module' && node.moduleType === 'class') {
                await showClassModuleReferences(node);
                return;
            }
            if (node.kind !== 'sub') { return; }
            const uri = node.isRemote && node.remoteId
                ? encodeRemoteModuleUri(node.remoteId, node.moduleName)
                : encodeModuleUri(node.filePath, node.moduleName);

            const doc = await vscode.workspace.openTextDocument(uri);
            await vscode.languages.setTextDocumentLanguage(doc, XLIDE_VBA_LANGUAGE_ID);
            const editor = await vscode.window.showTextDocument(doc, { preview: false });

            // Locate the procedure name on its declaration line so the reference
            // search starts on the identifier. The node label is "<kind> <name>"
            // (kind may be "Property Get" etc.), so the bare name is the last token.
            const procName = node.label.split(' ').pop() ?? '';
            let pos = new vscode.Position(Math.max(0, (node.line ?? 1) - 1), 0);
            if (procName && node.line !== undefined && node.line > 0) {
                const lineText = doc.lineAt(node.line - 1).text;
                const col = lineText.indexOf(procName);
                if (col >= 0) {
                    pos = new vscode.Position(node.line - 1, col);
                }
            }

            // Move the active editor's cursor onto the identifier so the
            // references command resolves the correct symbol, then trigger it.
            editor.selection = new vscode.Selection(pos, pos);
            editor.revealRange(
                new vscode.Range(pos, pos),
                vscode.TextEditorRevealType.InCenterIfOutsideViewport,
            );
            await vscode.commands.executeCommand('references-view.findReferences', uri, pos);
        }),

        registerXlideCommand('xlide.newModule', async (node: XlideNode) => {
            if (node?.kind !== 'xlsm') { return; }
            const name = await vscode.window.showInputBox({
                prompt: 'New module name',
                placeHolder: 'Module1',
                validateInput: validateVbaModuleName,
            });
            if (!name) { return; }

            const stub = `Option Explicit\r\n\r\nSub ${name}_Main()\r\n\r\nEnd Sub\r\n`;
            try {
                await writeWorkbookModule(moduleOps, {
                    filePath: node.filePath,
                    moduleName: name,
                    source: stub,
                });
                const summaryText = logChangeSummary(log, 'newModule', {
                    operation: 'Create module',
                    changed: [name],
                });
                recordWriteAudit({
                    command: 'xlide.newModule',
                    operation: 'create-module',
                    outcome: 'succeeded',
                    workbookPath: node.filePath,
                    moduleName: name,
                    summary: summaryText,
                });
                // Open the new module immediately
                const uri = encodeModuleUri(node.filePath, name);
                const doc = await vscode.workspace.openTextDocument(uri);
                await vscode.languages.setTextDocumentLanguage(doc, XLIDE_VBA_LANGUAGE_ID);
                await vscode.window.showTextDocument(doc, { preview: false });
            } catch (err) {
                recordWriteAudit({
                    command: 'xlide.newModule',
                    operation: 'create-module',
                    outcome: 'failed',
                    workbookPath: node.filePath,
                    moduleName: name,
                    summary: 'Create module: 0 changed, 1 failed',
                    error: err,
                });
                vscode.window.showErrorMessage(`XLIDE: Failed to create module: ${err}`);
            }
        }),

        // Add a new class module
        registerXlideCommand('xlide.newClassModule', async (node: XlideNode) => {
            if (node?.kind !== 'xlsm') { return; }
            const name = await vscode.window.showInputBox({
                prompt: 'New class module name',
                placeHolder: 'MyClass',
                validateInput: validateVbaModuleName,
            });
            if (!name) { return; }

            const stub = `Option Explicit\r\n\r\nPrivate Sub Class_Initialize()\r\n\r\nEnd Sub\r\n\r\nPrivate Sub Class_Terminate()\r\n\r\nEnd Sub\r\n`;
            try {
                await writeWorkbookModule(moduleOps, {
                    filePath: node.filePath,
                    moduleName: name,
                    source: stub,
                    kind: 'class',
                });
                const summaryText = logChangeSummary(log, 'newClassModule', {
                    operation: 'Create class module',
                    changed: [name],
                });
                recordWriteAudit({
                    command: 'xlide.newClassModule',
                    operation: 'create-class-module',
                    outcome: 'succeeded',
                    workbookPath: node.filePath,
                    moduleName: name,
                    summary: summaryText,
                });
                const uri = encodeModuleUri(node.filePath, name);
                const doc = await vscode.workspace.openTextDocument(uri);
                await vscode.languages.setTextDocumentLanguage(doc, XLIDE_VBA_LANGUAGE_ID);
                await vscode.window.showTextDocument(doc, { preview: false });
            } catch (err) {
                recordWriteAudit({
                    command: 'xlide.newClassModule',
                    operation: 'create-class-module',
                    outcome: 'failed',
                    workbookPath: node.filePath,
                    moduleName: name,
                    summary: 'Create class module: 0 changed, 1 failed',
                    error: err,
                });
                vscode.window.showErrorMessage(`XLIDE: Failed to create class module: ${err}`);
            }
        }),

        // Rename a module
        registerXlideCommand('xlide.renameModule', async (node: XlideNode) => {
            if (!node?.moduleName) { return; }
            const newName = await vscode.window.showInputBox({
                prompt: `Rename "${node.moduleName}" to`,
                value: node.moduleName,
                validateInput: validateVbaModuleName,
            });
            if (!newName || newName === node.moduleName) { return; }

            let moduleRenamed = false;
            try {
                if (node.moduleType === 'class') {
                    const modules = applyOpenDocumentSources(
                        await vbaIndex.getAllModules(node.filePath),
                        node.filePath,
                    );
                    const project = await buildVbaProjectIndexAsync(modules);
                    const byModule = new Map(modules.map((mod) => [mod.moduleName.toLowerCase(), mod]));
                    const definition = projectClassModuleDefinition(
                        project,
                        node.moduleName,
                        node.moduleName,
                    );
                    if (!definition) {
                        throw new Error(`"${node.moduleName}" is not a project-defined class module.`);
                    }
                    const references = projectClassReferenceEdit(
                        node.filePath,
                        byModule,
                        project,
                        node.moduleName,
                        definition,
                        newName,
                    );
                    await renameProjectClassModule(bridge, node.filePath, node.moduleName, newName);
                    moduleRenamed = true;
                    vbaIndex.invalidate(node.filePath);
                    if (references.count > 0) {
                        for (const uri of references.uris) {
                            const doc = await vscode.workspace.openTextDocument(uri);
                            await vscode.languages.setTextDocumentLanguage(doc, XLIDE_VBA_LANGUAGE_ID);
                        }
                        const applied = await vscode.workspace.applyEdit(references.edit);
                        if (!applied) {
                            throw new Error('VS Code did not apply the class reference edits.');
                        }
                    }
                } else {
                    const modules = applyOpenDocumentSources(
                        await vbaIndex.getAllModules(node.filePath),
                        node.filePath,
                    );
                    const project = await buildVbaProjectIndexAsync(modules);
                    const byModule = new Map(modules.map((mod) => [mod.moduleName.toLowerCase(), mod]));
                    const references = projectStandardModuleReferenceEdit(
                        node.filePath,
                        byModule,
                        project,
                        node.moduleName,
                        newName,
                    );
                    await renameProjectStandardModule(bridge, node.filePath, node.moduleName, newName);
                    moduleRenamed = true;
                    vbaIndex.invalidate(node.filePath);
                    if (references.count > 0) {
                        for (const uri of references.uris) {
                            const doc = await vscode.workspace.openTextDocument(uri);
                            await vscode.languages.setTextDocumentLanguage(doc, XLIDE_VBA_LANGUAGE_ID);
                        }
                        const applied = await vscode.workspace.applyEdit(references.edit);
                        if (!applied) {
                            throw new Error('VS Code did not apply the standard module reference edits.');
                        }
                    }
                }
                // Tell open editors the old module is gone and refresh workbook stats
                fsProvider.notifyFileChanged(encodeModuleUri(node.filePath, node.moduleName));
                const summaryText = logChangeSummary(log, 'renameModule', {
                    operation: 'Rename module',
                    changed: [`${node.moduleName} -> ${newName}`],
                });
                recordWriteAudit({
                    command: 'xlide.renameModule',
                    operation: 'rename-module',
                    outcome: 'succeeded',
                    workbookPath: node.filePath,
                    moduleName: newName,
                    summary: summaryText,
                });
            } catch (err) {
                const prefix = moduleRenamed
                    ? 'XLIDE: Module was renamed, but reference updates failed'
                    : 'XLIDE: Rename failed';
                recordWriteAudit({
                    command: 'xlide.renameModule',
                    operation: 'rename-module',
                    outcome: 'failed',
                    workbookPath: node.filePath,
                    moduleName: moduleRenamed ? newName : node.moduleName,
                    summary: moduleRenamed
                        ? 'Rename module: 1 changed, 1 failed'
                        : 'Rename module: 0 changed, 1 failed',
                    error: err,
                });
                vscode.window.showErrorMessage(`${prefix}: ${err}`);
            } finally {
                if (moduleRenamed) {
                    refreshWorkbookProjectState(moduleOps, node.filePath);
                }
            }
        }),

        // Delete a module (with confirmation)
        registerXlideCommand('xlide.deleteModule', async (node: XlideNode) => {
            if (!node?.moduleName) { return; }

            // Prevent deletion of document-type modules
            if (node.moduleType === 'document') {
                vscode.window.showWarningMessage(
                    `Cannot delete "${node.moduleName}" — document modules are protected.`,
                );
                return;
            }

            const choice = await vscode.window.showWarningMessage(
                `Delete module "${node.moduleName}" from "${path.basename(node.filePath)}"?`,
                { modal: true },
                'Delete',
            );
            if (choice !== 'Delete') { return; }

            try {
                await deleteWorkbookModule(moduleOps, {
                    filePath: node.filePath,
                    moduleName: node.moduleName,
                });
                const summaryText = logChangeSummary(log, 'deleteModule', {
                    operation: 'Delete module',
                    changed: [node.moduleName],
                });
                recordWriteAudit({
                    command: 'xlide.deleteModule',
                    operation: 'delete-module',
                    outcome: 'succeeded',
                    workbookPath: node.filePath,
                    moduleName: node.moduleName,
                    summary: summaryText,
                });
            } catch (err) {
                recordWriteAudit({
                    command: 'xlide.deleteModule',
                    operation: 'delete-module',
                    outcome: 'failed',
                    workbookPath: node.filePath,
                    moduleName: node.moduleName,
                    summary: 'Delete module: 0 changed, 1 failed',
                    error: err,
                });
                vscode.window.showErrorMessage(`XLIDE: Delete failed: ${err}`);
            }
        }),

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

        ...registerSupportBundleCommands(deps),

        // DEV: smoke test — verifies listModules + readModule against a workspace workbook
        registerXlideCommand('xlide.dev.smoke', async () => {
            log('[smoke] Starting smoke test...');

            const uris = (await vscode.workspace.findFiles('**/*.{xlsm,xlsb,xlam}',
                '{**/node_modules/**,**/.venv/**,**/venv/**}'))
                .filter(u => !path.basename(u.fsPath).startsWith('~$'));

            if (uris.length === 0) {
                vscode.window.showErrorMessage('XLIDE Smoke: No workbook found in the workspace.');
                return;
            }

            let workbookPath: string;
            if (uris.length === 1) {
                workbookPath = uris[0].fsPath;
            } else {
                const pick = await vscode.window.showQuickPick(
                    uris.map(u => ({ label: path.basename(u.fsPath), description: u.fsPath, fsPath: u.fsPath })),
                    { title: 'XLIDE Smoke Test: pick a workbook' },
                );
                if (!pick) { return; }
                workbookPath = pick.fsPath;
            }

            log(`[smoke] Workbook: ${workbookPath}`);

            await vscode.window.withProgress(
                { location: vscode.ProgressLocation.Notification, title: 'XLIDE: Running smoke test...', cancellable: false },
                async () => {
                    try {
                        // Step 1: listModules
                        const modules = await bridge.call<Array<{ name: string; type: string }>>(
                            'listModules', { path: workbookPath },
                        );
                        log(`[smoke] listModules OK — ${modules.length} module(s): ${modules.map(m => m.name).join(', ')}`);

                        if (modules.length === 0) {
                            vscode.window.showWarningMessage('XLIDE Smoke: workbook has no VBA modules.');
                            return;
                        }

                        // Step 2: readModule (prefer a non-document module)
                        const target = modules.find(m => m.type !== 'document') ?? modules[0];
                        const source = await bridge.call<string>(
                            'readModule', { path: workbookPath, module: target.name, full: false },
                        );
                        log(`[smoke] readModule "${target.name}" OK — ${source.length} chars`);

                        log('[smoke] All checks passed.');
                        void vscode.window.showInformationMessage(
                            `XLIDE Smoke: OK — ${modules.length} modules, read "${target.name}" (${source.length} chars). See XLIDE Output for details.`,
                        );
                    } catch (err) {
                        const msg = errorMessage(err);
                        log(`[smoke] FAILED: ${msg}`);
                        vscode.window.showErrorMessage(`XLIDE Smoke FAILED: ${msg}`);
                    }
                },
            );
        }),

        ...registerAnalysisCommands(deps),
        ...registerVbaTestCommands(deps),

        // Create a new, empty macro-enabled workbook
        registerXlideCommand('xlide.newWorkbook', async () => {
            const defaultDir = vscode.workspace.workspaceFolders?.[0]?.uri;
            const target = await vscode.window.showSaveDialog({
                title: 'XLIDE: New Macro-Enabled Workbook',
                defaultUri: defaultDir ? vscode.Uri.joinPath(defaultDir, 'NewWorkbook.xlsm') : undefined,
                filters: { 'Macro-Enabled Workbook': ['xlsm', 'xlsb'] },
            });
            if (!target) { return; }
            const filePath = target.fsPath;
            const name = path.basename(filePath);
            await vscode.window.withProgress(
                { location: vscode.ProgressLocation.Notification, title: `XLIDE: Creating "${name}"...`, cancellable: false },
                async () => {
                    await bridge.call<{ ok: boolean; path: string }>('createWorkbook', { path: filePath });
                    log(`[newWorkbook] Created "${filePath}"`);
                    explorer.refresh();
                    void vscode.window.showInformationMessage(`XLIDE: Created "${name}".`);
                },
            );
        }, { errorPrefix: 'Failed to create workbook', logTag: 'newWorkbook', log }),

        // Open the workbook in Excel (editable)
        registerXlideCommand('xlide.openWorkbook', async (node: XlideNode) => {
            const filePath = resolveWorkbookPath(node);
            if (!filePath) { return; }
            try {
                const attachToRunning = shouldAttachToRunningExcel();
                log(`[openWorkbook] Requested for: ${filePath}`);
                if (process.platform === 'win32') {
                    runWindowsExcel(filePath, attachToRunning, false);
                } else if (process.platform === 'darwin') {
                    cp.spawn('open', ['-a', 'Microsoft Excel', filePath]);
                } else {
                    cp.spawn('libreoffice', ['--calc', '--norestore', filePath]);
                }
            } catch (err) {
                vscode.window.showErrorMessage(`Failed to open workbook: ${err}`);
            }
        }),

        // Open the workbook in Excel (read-only)
        registerXlideCommand('xlide.openWorkbookReadOnly', async (node: XlideNode) => {
            const filePath = resolveWorkbookPath(node);
            if (!filePath) { return; }
            try {
                const attachToRunning = shouldAttachToRunningExcel();
                log(`[openWorkbookReadOnly] Requested for: ${filePath}`);
                if (process.platform === 'win32') {
                    runWindowsExcel(filePath, attachToRunning, true);
                } else if (process.platform === 'darwin') {
                    cp.spawn('open', ['-a', 'Microsoft Excel', filePath]);
                } else {
                    cp.spawn('libreoffice', ['--calc', '--norestore', '--view', filePath]);
                }
            } catch (err) {
                vscode.window.showErrorMessage(`Failed to open workbook: ${err}`);
            }
        }),

        // Detect the Sub/Function at the cursor and open the workbook, then guide to run it
        registerXlideCommand('xlide.runMacroAtCursor', async () => {
            const editor = vscode.window.activeTextEditor;
            if (!editor || editor.document.uri.scheme !== XLIDE_SCHEME) {
                vscode.window.showWarningMessage('XLIDE: Open a VBA module to run a macro.');
                return;
            }

            try {
                // Persist any in-editor changes first so the macro that runs
                // reflects the current source rather than the last-saved version.
                if (editor.document.isDirty) {
                    await editor.document.save();
                }

                // Decode the URI to get filePath and moduleName
                const { xlsmPath, moduleName } = decodeModuleUri(editor.document.uri);
                log(`[runMacro] Requested from module: ${moduleName} in ${xlsmPath}`);

                // Find which procedure the cursor is in (parser-based, so
                // Friend/Global/Static modifiers are recognized too).
                const currentProc = procedureNameAtCursor(editor);
                if (!currentProc) {
                    vscode.window.showWarningMessage('XLIDE: Cursor is not inside a Sub or Function.');
                    return;
                }

                // Open the workbook read-only
                if (process.platform === 'win32') {
                    const attachToRunning = shouldAttachToRunningExcel();
                    log(`[runMacro] attachToRunningExcel=${attachToRunning}`);
                    await runWorkbookMacroReadOnly(xlsmPath, `${moduleName}.${currentProc}`, { attachToRunning }, log);
                } else if (process.platform === 'darwin') {
                    cp.spawn('open', ['-a', 'Microsoft Excel', xlsmPath]);
                    vscode.window.showInformationMessage(
                        `Workbook opened. Run macro: ${moduleName}.${currentProc}`,
                    );
                } else {
                    cp.spawn('libreoffice', ['--calc', '--norestore', '--view', xlsmPath]);
                    vscode.window.showInformationMessage(
                        `Workbook opened. Run macro manually: ${moduleName}.${currentProc}`,
                    );
                }
            } catch (err) {
                showRunMacroFailure(err);
            }
        }),
    ];
}
