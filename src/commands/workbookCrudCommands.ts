import * as vscode from 'vscode';
import { takeRenameForUndo } from '../vbaRenameHistory';
import * as path from 'path';
import {
    encodeModuleUri,
    isWorkbookLockedError,
    reportWorkbookLocked,
    XLIDE_VBA_LANGUAGE_ID,
} from '../xlideFileSystem';
import type { WorkbookEngine } from '../workbookEngine';
import { errorMessage } from '../util/errors';
import { applyOpenDocumentSources } from '../vbaOpenDocuments';
import { validateVbaModuleName } from '../vbaSourceScan';
import { projectClassModuleDefinition } from '../vbaNavigation';
import { buildVbaProjectIndexAsync } from '../vbaProjectAnalysis';
import {
    projectClassReferenceEdit,
    renameProjectClassModule,
} from '../vbaClassRename';
import {
    projectStandardModuleReferenceEdit,
    renameProjectStandardModule,
} from '../vbaStandardModuleRename';
import { recordXlideWriteAuditEvent as recordWriteAudit } from '../xlideWriteAudit';
import {
    deleteWorkbookModule,
    refreshWorkbookProjectState,
    writeWorkbookModule,
} from '../workbookModuleOperations';
import { registerXlideCommand } from '../xlideCommandRegistration';
import type { XlideNode } from '../xlsmExplorer';
import {
    logChangeSummary,
    type CommandDeps,
} from './shared';

/** Best-effort lowercased set of the workbook's existing module names. */
async function existingModuleNamesLower(bridge: WorkbookEngine, filePath: string): Promise<Set<string>> {
    try {
        const modules = await bridge.call<Array<{ name: string }>>('listModules', { path: filePath });
        return new Set(modules.map((m) => m.name.toLowerCase()));
    } catch {
        // The workbook may be locked/unavailable; the write itself will surface
        // that. Skip the duplicate check rather than blocking the prompt.
        return new Set();
    }
}

/**
 * Prompts for a new module name, rejecting both invalid syntax and names that
 * already exist in the workbook (an existing name would silently overwrite that
 * module's code). Returns undefined if the user cancels.
 */
async function promptForNewModuleName(
    bridge: WorkbookEngine,
    filePath: string,
    options: { prompt: string; placeHolder: string },
): Promise<string | undefined> {
    const existing = await existingModuleNamesLower(bridge, filePath);
    return vscode.window.showInputBox({
        prompt: options.prompt,
        placeHolder: options.placeHolder,
        validateInput: (value) => {
            const syntax = validateVbaModuleName(value);
            if (syntax) { return syntax; }
            if (existing.has(value.trim().toLowerCase())) {
                return `A module named "${value.trim()}" already exists in this workbook.`;
            }
            return undefined;
        },
    });
}

/** Surfaces a workbook write failure, preferring the friendly "open in Excel" notice. */
function surfaceWorkbookWriteError(filePath: string, err: unknown, fallbackPrefix: string): void {
    if (isWorkbookLockedError(errorMessage(err))) {
        reportWorkbookLocked(filePath, 'write');
    } else {
        void vscode.window.showErrorMessage(`${fallbackPrefix}: ${err}`);
    }
}

export function registerWorkbookCrudCommands(deps: CommandDeps): vscode.Disposable[] {
    const { bridge, explorer, fsProvider, out, vbaIndex } = deps;

    function log(msg: string): void {
        out.appendLine(msg);
    }

    return [
        // Put the last rename back as one operation (issue #9 rule 10).
        // A rename edits several modules; an editor's undo stack is per
        // document, so undoing in one file would leave the rest renamed.
        registerXlideCommand('xlide.undoRename', async () => {
            const snapshot = takeRenameForUndo();
            if (!snapshot) {
                void vscode.window.showInformationMessage(
                    'XLIDE: there is no rename to undo. Only the most recent rename can be put '
                    + 'back, and only until something else writes to the workbook.',
                );
                return;
            }
            try {
                for (const image of snapshot.modules) {
                    await writeWorkbookModule(
                        { bridge, explorer, fsProvider, vbaIndex },
                        { filePath: snapshot.workbookPath, moduleName: image.moduleName, source: image.before },
                    );
                }
                vbaIndex.invalidate(snapshot.workbookPath);
                const summaryText = logChangeSummary(log, 'undoRename', {
                    operation: 'Undo rename',
                    changed: snapshot.modules.map((image) => image.moduleName),
                });
                void vscode.window.showInformationMessage(`XLIDE: ${summaryText}`);
            } catch (err) {
                void vscode.window.showErrorMessage(
                    `XLIDE: could not undo the rename - ${err instanceof Error ? err.message : String(err)}`,
                );
            }
        }),

        // Create a new, empty macro-enabled file for any supported host
        registerXlideCommand('xlide.newWorkbook', async () => {
            const defaultDir = vscode.workspace.workspaceFolders?.[0]?.uri;
            const target = await vscode.window.showSaveDialog({
                title: 'XLIDE: New Macro-Enabled File',
                // A host-neutral base name: the dialog keeps it when the user
                // switches the type filter, so "NewWorkbook.docm" must never
                // be the default a Word document is born with.
                defaultUri: defaultDir ? vscode.Uri.joinPath(defaultDir, 'NewFile.xlsm') : undefined,
                // One filter per file kind: the dialog auto-appends only the
                // FIRST extension of the selected filter, so a bundled
                // "docm;dotm" filter could never produce a .dotm without the
                // user typing the extension by hand.
                filters: {
                    'Excel Macro-Enabled Workbook': ['xlsm'],
                    'Excel Binary Workbook': ['xlsb'],
                    'Excel Add-In': ['xlam'],
                    'Excel Macro-Enabled Template': ['xltm'],
                    'Word Macro-Enabled Document': ['docm'],
                    'Word Macro-Enabled Template': ['dotm'],
                    'PowerPoint Macro-Enabled Presentation': ['pptm'],
                    'PowerPoint Macro-Enabled Template': ['potm'],
                },
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
        }, { errorPrefix: 'Failed to create file', logTag: 'newWorkbook', log }),

        registerXlideCommand('xlide.newModule', async (node: XlideNode) => {
            if (node?.kind !== 'xlsm') { return; }
            const name = await promptForNewModuleName(bridge, node.filePath, {
                prompt: 'New module name',
                placeHolder: 'Module1',
            });
            if (!name) { return; }

            const stub = `Option Explicit\r\n\r\nSub ${name}_Main()\r\n\r\nEnd Sub\r\n`;
            try {
                await writeWorkbookModule(deps, {
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
                surfaceWorkbookWriteError(node.filePath, err, 'XLIDE: Failed to create module');
            }
        }),

        // Add a new class module
        registerXlideCommand('xlide.newClassModule', async (node: XlideNode) => {
            if (node?.kind !== 'xlsm') { return; }
            const name = await promptForNewModuleName(bridge, node.filePath, {
                prompt: 'New class module name',
                placeHolder: 'MyClass',
            });
            if (!name) { return; }

            const stub = `Option Explicit\r\n\r\nPrivate Sub Class_Initialize()\r\n\r\nEnd Sub\r\n\r\nPrivate Sub Class_Terminate()\r\n\r\nEnd Sub\r\n`;
            try {
                await writeWorkbookModule(deps, {
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
                surfaceWorkbookWriteError(node.filePath, err, 'XLIDE: Failed to create class module');
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
                surfaceWorkbookWriteError(node.filePath, err, prefix);
            } finally {
                if (moduleRenamed) {
                    refreshWorkbookProjectState(deps, node.filePath);
                }
            }
        }),

        // Delete a module (with confirmation)
        registerXlideCommand('xlide.deleteModule', async (node: XlideNode) => {
            if (!node?.moduleName) { return; }

            // Prevent deletion of document-type modules
            if (node.moduleType === 'document') {
                vscode.window.showWarningMessage(
                    `Cannot delete "${node.moduleName}": document modules are protected.`,
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
                await deleteWorkbookModule(deps, {
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
                surfaceWorkbookWriteError(node.filePath, err, 'XLIDE: Delete failed');
            }
        }),
    ];
}
