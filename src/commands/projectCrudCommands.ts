import * as vscode from 'vscode';
import { takeRenameForUndo } from '../vbaRenameHistory';
import * as path from 'path';
import {
    encodeFormMarkupUri,
    isProjectLockedError,
    reportProjectLocked,
    XLIDE_VBA_LANGUAGE_ID,
} from '../xlideFileSystem';
import type { ProjectEngine } from '../projectEngine';
import { errorMessage } from '../util/errors';
import { applyOpenDocumentSources } from '../vbaOpenDocuments';
import { validateVbaModuleName } from '../vbaSourceScan';
import { projectClassModuleDefinition } from '../vbaNavigation';
import { buildVbaProjectIndexAsync } from '../vbaProjectAnalysis';
import { projectClassReferenceEdit } from '../vbaClassRename';
import { projectStandardModuleReferenceEdit } from '../vbaStandardModuleRename';
import { projectUserFormReferenceEdit } from '../vbaUserFormRename';
import { recordXlideWriteAuditEvent as recordWriteAudit } from '../xlideWriteAudit';
import {
    deleteProjectModule,
    refreshProjectState,
    renameProjectModule,
    writeProjectModule,
} from '../projectModuleOperations';
import { registerXlideCommand } from '../xlideCommandRegistration';
import { runWriteWithHostCoordination } from '../officeWriteCoordinator';
import type { XlideNode } from '../projectExplorer';
import {
    logChangeSummary,
    type CommandDeps,
    outputLogger,
    openModuleDocument,
} from './shared';

/** Best-effort lowercased set of the project's existing module names. */
async function existingModuleNamesLower(bridge: ProjectEngine, filePath: string): Promise<Set<string>> {
    try {
        const modules = await bridge.call<Array<{ name: string }>>('listModules', { path: filePath });
        return new Set(modules.map((m) => m.name.toLowerCase()));
    } catch {
        // The project may be locked/unavailable; the write itself will surface
        // that. Skip the duplicate check rather than blocking the prompt.
        return new Set();
    }
}

/**
 * Prompts for a new module name, rejecting both invalid syntax and names that
 * already exist in the project (an existing name would silently overwrite that
 * module's code). Returns undefined if the user cancels.
 */
async function promptForNewModuleName(
    bridge: ProjectEngine,
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
                return `A module named "${value.trim()}" already exists in this project.`;
            }
            return undefined;
        },
    });
}

/** Surfaces a project write failure, preferring the friendly "open in Excel" notice. */
function surfaceProjectWriteError(filePath: string, err: unknown, fallbackPrefix: string): void {
    if (isProjectLockedError(errorMessage(err))) {
        reportProjectLocked(filePath, 'write');
    } else {
        void vscode.window.showErrorMessage(`${fallbackPrefix}: ${err}`);
    }
}

/**
 * Module kinds that have a design to open: a UserForm, an Access form or
 * report, and the VB6 designers, which carry their own header instead.
 */
const DESIGNABLE_MODULE_TYPES: ReadonlySet<string> = new Set([
    'userform', 'accessform', 'accessreport',
]);

export function registerProjectCrudCommands(deps: CommandDeps): vscode.Disposable[] {
    const { bridge, explorer, fsProvider, out, vbaIndex } = deps;

    const log = outputLogger(out);

    /** What each designer object is called where the user can see it. */
    const DESIGNER_WORDS = {
        userform: { title: 'UserForm', command: 'xlide.newUserForm', placeHolder: 'FrmMain' },
        form: { title: 'form', command: 'xlide.newAccessForm', placeHolder: 'Orders' },
        report: { title: 'report', command: 'xlide.newAccessReport', placeHolder: 'Invoices' },
    } as const;

    /**
     * Create a designer object and open its markup: a UserForm in the project,
     * or an Access form or report in the database. The engine decides which
     * from the container; the kind here is what the user asked for and what
     * the messages say.
     */
    async function newDesigner(
        node: XlideNode,
        kind: keyof typeof DESIGNER_WORDS,
    ): Promise<void> {
        if (node?.kind !== 'project') { return; }
        const words = DESIGNER_WORDS[kind];
        const name = await promptForNewModuleName(bridge, node.filePath, {
            prompt: `New ${words.title} name`,
            placeHolder: words.placeHolder,
        });
        if (!name) { return; }
        try {
            const created = await runWriteWithHostCoordination(node.filePath, () =>
                bridge.call('addForm', {
                    path: node.filePath,
                    module: name,
                    source: 'Option Explicit\r\n',
                    ...(kind === 'userform' ? {} : { kind }),
                })) as { moduleName?: string } | undefined;
            const opened = created?.moduleName ?? name;
            const summaryText = logChangeSummary(log, words.command.replace('xlide.', ''), {
                operation: `Create ${words.title}`,
                changed: [opened],
            });
            recordWriteAudit({
                command: words.command,
                operation: `create-${kind}`,
                outcome: 'succeeded',
                projectPath: node.filePath,
                moduleName: opened,
                summary: summaryText,
            });
            explorer.refresh();
            const doc = await vscode.workspace.openTextDocument(
                encodeFormMarkupUri(node.filePath, opened),
            );
            await vscode.languages.setTextDocumentLanguage(doc, 'xml');
            await vscode.window.showTextDocument(doc, { preview: false });
        } catch (err) {
            recordWriteAudit({
                command: words.command,
                operation: `create-${kind}`,
                outcome: 'failed',
                projectPath: node.filePath,
                moduleName: name,
                summary: `Create ${words.title}: 0 changed, 1 failed`,
                error: err,
            });
            surfaceProjectWriteError(
                node.filePath, err, `XLIDE: Failed to create the ${words.title}`,
            );
        }
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
                    + 'back, and only until something else writes to the project.',
                );
                return;
            }
            try {
                for (const image of snapshot.modules) {
                    await writeProjectModule(
                        { bridge, explorer, fsProvider, vbaIndex },
                        { filePath: snapshot.projectPath, moduleName: image.moduleName, source: image.before },
                    );
                }
                vbaIndex.invalidate(snapshot.projectPath);
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
        registerXlideCommand('xlide.newProject', async () => {
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
                    'Access Database': ['accdb'],
                    'Access Add-In': ['accda'],
                    'Access 2002-2003 Database': ['mdb'],
                    'Access 2002-2003 Add-In': ['mda'],
                },
            });
            if (!target) { return; }
            const filePath = target.fsPath;
            const name = path.basename(filePath);
            await vscode.window.withProgress(
                { location: vscode.ProgressLocation.Notification, title: `XLIDE: Creating "${name}"...`, cancellable: false },
                async () => {
                    await bridge.call<{ ok: boolean; path: string }>('createProject', { path: filePath });
                    log(`[newProject] Created "${filePath}"`);
                    explorer.refresh();
                    void vscode.window.showInformationMessage(`XLIDE: Created "${name}".`);
                },
            );
        }, { errorPrefix: 'Failed to create file', logTag: 'newProject', log }),

        registerXlideCommand('xlide.newModule', async (node: XlideNode) => {
            if (node?.kind !== 'project') { return; }
            const name = await promptForNewModuleName(bridge, node.filePath, {
                prompt: 'New module name',
                placeHolder: 'Module1',
            });
            if (!name) { return; }

            const stub = `Option Explicit\r\n\r\nSub ${name}_Main()\r\n\r\nEnd Sub\r\n`;
            try {
                await writeProjectModule(deps, {
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
                    projectPath: node.filePath,
                    moduleName: name,
                    summary: summaryText,
                });
                // Open the new module immediately
                const doc = await openModuleDocument(node.filePath, name);
                await vscode.window.showTextDocument(doc, { preview: false });
            } catch (err) {
                recordWriteAudit({
                    command: 'xlide.newModule',
                    operation: 'create-module',
                    outcome: 'failed',
                    projectPath: node.filePath,
                    moduleName: name,
                    summary: 'Create module: 0 changed, 1 failed',
                    error: err,
                });
                surfaceProjectWriteError(node.filePath, err, 'XLIDE: Failed to create module');
            }
        }),

        // Add a new class module
        registerXlideCommand('xlide.newClassModule', async (node: XlideNode) => {
            if (node?.kind !== 'project') { return; }
            const name = await promptForNewModuleName(bridge, node.filePath, {
                prompt: 'New class module name',
                placeHolder: 'MyClass',
            });
            if (!name) { return; }

            const stub = `Option Explicit\r\n\r\nPrivate Sub Class_Initialize()\r\n\r\nEnd Sub\r\n\r\nPrivate Sub Class_Terminate()\r\n\r\nEnd Sub\r\n`;
            try {
                await writeProjectModule(deps, {
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
                    projectPath: node.filePath,
                    moduleName: name,
                    summary: summaryText,
                });
                const doc = await openModuleDocument(node.filePath, name);
                await vscode.window.showTextDocument(doc, { preview: false });
            } catch (err) {
                recordWriteAudit({
                    command: 'xlide.newClassModule',
                    operation: 'create-class-module',
                    outcome: 'failed',
                    projectPath: node.filePath,
                    moduleName: name,
                    summary: 'Create class module: 0 changed, 1 failed',
                    error: err,
                });
                surfaceProjectWriteError(node.filePath, err, 'XLIDE: Failed to create class module');
            }
        }),

        // Add a new UserForm: the module with its code-behind and a designer
        // storage authored natively - no Office application is involved.
        registerXlideCommand('xlide.newUserForm', (node: XlideNode) => newDesigner(node, 'userform')),
        // An Access database keeps its designer objects in the database rather
        // than in the project, and has two kinds of them.
        registerXlideCommand('xlide.newAccessForm', (node: XlideNode) => newDesigner(node, 'form')),
        registerXlideCommand('xlide.newAccessReport', (node: XlideNode) => newDesigner(node, 'report')),

        // Open a form's markup projection beside its code-behind.
        registerXlideCommand('xlide.openFormMarkup', async (node: XlideNode) => {
            if (node?.kind !== 'module' || !node.moduleName
                || !DESIGNABLE_MODULE_TYPES.has(node.moduleType ?? '')) { return; }
            try {
                const doc = await vscode.workspace.openTextDocument(encodeFormMarkupUri(node.filePath, node.moduleName));
                await vscode.languages.setTextDocumentLanguage(doc, 'xml');
                await vscode.window.showTextDocument(doc, { preview: false });
            } catch (err) {
                void vscode.window.showErrorMessage(`XLIDE: ${errorMessage(err)}`);
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
            let renamedTo = newName;
            try {
                const modules = applyOpenDocumentSources(
                    await vbaIndex.getAllModules(node.filePath),
                    node.filePath,
                );
                const project = await buildVbaProjectIndexAsync(modules);
                const byModule = new Map(modules.map((mod) => [mod.moduleName.toLowerCase(), mod]));
                // The shared operation coordinates with Excel, carries a pending
                // agent review to the new name and tells open editors the old
                // module is gone; project state refreshes once, in `finally`.
                const renameRequest = { filePath: node.filePath, moduleName: node.moduleName, newName };
                const rename = async (): Promise<void> => {
                    const result = await renameProjectModule(deps, renameRequest, { refreshProjectState: false });
                    // An Access form's module keeps its prefix: `Form_Customers`.
                    renamedTo = result.moduleName ?? newName;
                    moduleRenamed = true;
                    vbaIndex.invalidate(node.filePath);
                };
                if (node.moduleType === 'class') {
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
                    await rename();
                    await applyReferenceEdit(references, 'class');
                } else if (node.moduleType === 'userform') {
                    // A form's name is a type and its default instance:
                    // `Dim f As UserForm1` and `UserForm1.Show` both follow.
                    const references = projectUserFormReferenceEdit(
                        node.filePath,
                        byModule,
                        project,
                        node.moduleName,
                        newName,
                    );
                    await rename();
                    await applyReferenceEdit(references, 'form');
                } else {
                    const references = projectStandardModuleReferenceEdit(
                        node.filePath,
                        byModule,
                        project,
                        node.moduleName,
                        newName,
                    );
                    await rename();
                    await applyReferenceEdit(references, 'standard module');
                }
                const summaryText = logChangeSummary(log, 'renameModule', {
                    operation: 'Rename module',
                    changed: [`${node.moduleName} -> ${renamedTo}`],
                });
                recordWriteAudit({
                    command: 'xlide.renameModule',
                    operation: 'rename-module',
                    outcome: 'succeeded',
                    projectPath: node.filePath,
                    moduleName: renamedTo,
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
                    projectPath: node.filePath,
                    moduleName: moduleRenamed ? renamedTo : node.moduleName,
                    summary: moduleRenamed
                        ? 'Rename module: 1 changed, 1 failed'
                        : 'Rename module: 0 changed, 1 failed',
                    error: err,
                });
                surfaceProjectWriteError(node.filePath, err, prefix);
            } finally {
                if (moduleRenamed) {
                    refreshProjectState(deps, node.filePath);
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
                await deleteProjectModule(deps, {
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
                    projectPath: node.filePath,
                    moduleName: node.moduleName,
                    summary: summaryText,
                });
            } catch (err) {
                recordWriteAudit({
                    command: 'xlide.deleteModule',
                    operation: 'delete-module',
                    outcome: 'failed',
                    projectPath: node.filePath,
                    moduleName: node.moduleName,
                    summary: 'Delete module: 0 changed, 1 failed',
                    error: err,
                });
                surfaceProjectWriteError(node.filePath, err, 'XLIDE: Delete failed');
            }
        }),
    ];
}

/** Opens every module the rename touches as VBA, so the edit lands on the right language, then applies it. */
async function applyReferenceEdit(
    references: { count: number; uris: readonly vscode.Uri[]; edit: vscode.WorkspaceEdit },
    what: string,
): Promise<void> {
    if (references.count === 0) { return; }
    for (const uri of references.uris) {
        const doc = await vscode.workspace.openTextDocument(uri);
        await vscode.languages.setTextDocumentLanguage(doc, XLIDE_VBA_LANGUAGE_ID);
    }
    const applied = await vscode.workspace.applyEdit(references.edit);
    if (!applied) {
        throw new Error(`VS Code did not apply the ${what} reference edits.`);
    }
}
