import * as vscode from 'vscode';
import * as path from 'path';
import * as cp from 'child_process';
import {
    encodeModuleUri,
    decodeModuleUri,
    XLIDE_SCHEME,
    XLIDE_VBA_LANGUAGE_ID,
} from '../xlideFileSystem';
import { xlideAttachToRunningExcelFromConfig } from '../globalSettings';
import { containerAppNameForPath, containerHostForPath } from '../macroContainerUi';
import { registerXlideCommand } from '../xlideCommandRegistration';
import type { XlideNode } from '../projectExplorer';
import { errorMessage } from '../util/errors';
import {
    ExcelMacroError,
    openWorkbookInExcel,
    runHostFileMacroReadOnly,
    runWorkbookMacroReadOnly,
} from '../excelLauncher';
import {
    closeWorkbookInExcel,
    markWorkbookOpenedByXlide,
    resolveExcelCoordinationSettings,
    shouldAttemptClose,
    withWorkbookReopenSuppressed,
} from '../excelWorkbookCoordinator';
import {
    procedureAtCursor,
    requiredParameterNames,
    resolveProjectPath,
    type CommandDeps,
} from './shared';

export function registerMiscCommands(deps: CommandDeps): vscode.Disposable[] {
    const { bridge, explorer, out } = deps;

    function log(msg: string): void {
        out.appendLine(msg);
    }

    function shouldAttachToRunningExcel(): boolean {
        return xlideAttachToRunningExcelFromConfig(vscode.workspace.getConfiguration('xlide')).value;
    }

    function showRunMacroFailure(err: unknown, appName = 'Excel'): void {
        if (err instanceof ExcelMacroError &&
            (err.code === 'REOPEN_BLOCKED' || err.code === 'REOPEN_FAILED')) {
            void vscode.window.showWarningMessage(`XLIDE: ${err.message}`);
            return;
        }
        const message = errorMessage(err);
        // The host was busy and kept rejecting the COM call even after XLIDE's
        // retries (RPC_E_CALL_REJECTED / RETRYLATER) - almost always a modal
        // dialog left open, such as a MsgBox from a previous run.
        if (/rejected by callee|RPC_E_CALL_REJECTED|0x80010001|RETRYLATER|0x8001010A/i.test(message)) {
            void vscode.window.showWarningMessage(
                `XLIDE: ${appName} is busy, so the macro could not run. A dialog may be open in ${appName} `
                + '(for example a MsgBox from a previous run); close it, then press F5 again.',
            );
            return;
        }
        void vscode.window.showErrorMessage(`XLIDE: Failed to run macro: ${message}`);
    }

    // Windows COM-based Excel launch; the script lives in excelLauncher.ts.
    function runWindowsExcel(filePath: string, attachToRunning: boolean, readOnly: boolean): void {
        // Remember XLIDE opened this workbook so closeTracked coordination can
        // later close it without touching projects the user opened manually.
        markWorkbookOpenedByXlide(filePath);
        void openWorkbookInExcel(filePath, { attachToRunning, readOnly }, log).catch((err: Error) => {
            void vscode.window.showErrorMessage(`XLIDE: Open Workbook failed: ${err.message}`);
        });
    }

    async function showClassModuleReferences(node: XlideNode): Promise<void> {
        if (!node.moduleName || !node.filePath) { return; }
        const originUri = encodeModuleUri(node.filePath, node.moduleName);
        const originDoc = await vscode.workspace.openTextDocument(originUri);
        await vscode.languages.setTextDocumentLanguage(originDoc, XLIDE_VBA_LANGUAGE_ID);
        const editor = await vscode.window.showTextDocument(originDoc, { preview: false });
        const origin = new vscode.Position(0, 0);
        editor.selection = new vscode.Selection(origin, origin);
        await vscode.commands.executeCommand('references-view.findReferences', originUri, origin);
    }

    /**
     * The editor for a tree node's module: the project module's virtual
     * document, or - when the container's modules are files, a VB6 project -
     * the file itself, which already carries the VBA language by extension.
     */
    const showModuleEditor = async (node: XlideNode): Promise<vscode.TextEditor> => {
        if (node.moduleFilePath) {
            const doc = await vscode.workspace.openTextDocument(vscode.Uri.file(node.moduleFilePath));
            return vscode.window.showTextDocument(doc, { preview: false });
        }
        const uri = encodeModuleUri(node.filePath, node.moduleName ?? '');
        const doc = await vscode.workspace.openTextDocument(uri);
        await vscode.languages.setTextDocumentLanguage(doc, XLIDE_VBA_LANGUAGE_ID);
        return vscode.window.showTextDocument(doc, { preview: false });
    };

    return [
        registerXlideCommand('xlide.refreshExplorer', () => {
            explorer.refresh();
        }),

        // In-tree "Load failed - click to retry" placeholder (e.g. after Excel
        // briefly held the project file). Retries just the failed listing
        // instead of collapsing the whole tree with a full refresh.
        registerXlideCommand('xlide.retryExplorerLoad', (node: XlideNode) => {
            explorer.retryLoad(node);
        }),

        // Open a module (or navigate to a sub's line inside one)
        registerXlideCommand('xlide.openModule', async (node: XlideNode) => {
            if (!node?.moduleName) { return; }
            const editor = await showModuleEditor(node);

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
            const editor = await showModuleEditor(node);
            const doc = editor.document;

            // Locate the procedure name on its declaration line so the reference
            // search starts on the identifier. The node label is "<kind> <name>"
            // (kind may be "Property Get" etc.), so the bare name is the last token.
            const procName = node.label.split(' ').pop() ?? '';
            let pos = new vscode.Position(Math.max(0, (node.line ?? 1) - 1), 0);
            if (procName && node.line !== undefined && node.line > 0) {
                const lineText = doc.lineAt(node.line - 1).text;
                // Whole-word match so a short proc name that is a substring of a
                // preceding token (e.g. "i" inside "Public") does not mis-anchor.
                const escaped = procName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                const match = new RegExp(`\\b${escaped}\\b`).exec(lineText);
                const col = match ? match.index : lineText.indexOf(procName);
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
            await vscode.commands.executeCommand('references-view.findReferences', doc.uri, pos);
        }),

        // DEV: smoke test - verifies listModules + readModule against a workspace project
        registerXlideCommand('xlide.dev.smoke', async () => {
            log('[smoke] Starting smoke test...');

            const uris = (await vscode.workspace.findFiles('**/*.{xlsm,xlsb,xlam}',
                '{**/node_modules/**,**/.venv/**,**/venv/**}'))
                .filter(u => !path.basename(u.fsPath).startsWith('~$'));

            if (uris.length === 0) {
                vscode.window.showErrorMessage('XLIDE Smoke: No project found in the workspace.');
                return;
            }

            let projectPath: string;
            if (uris.length === 1) {
                projectPath = uris[0].fsPath;
            } else {
                const pick = await vscode.window.showQuickPick(
                    uris.map(u => ({ label: path.basename(u.fsPath), description: u.fsPath, fsPath: u.fsPath })),
                    { title: 'XLIDE Smoke Test: pick a project' },
                );
                if (!pick) { return; }
                projectPath = pick.fsPath;
            }

            log(`[smoke] Workbook: ${projectPath}`);

            await vscode.window.withProgress(
                { location: vscode.ProgressLocation.Notification, title: 'XLIDE: Running smoke test...', cancellable: false },
                async () => {
                    try {
                        // Step 1: listModules
                        const modules = await bridge.call<Array<{ name: string; type: string }>>(
                            'listModules', { path: projectPath },
                        );
                        log(`[smoke] listModules OK - ${modules.length} module(s): ${modules.map(m => m.name).join(', ')}`);

                        if (modules.length === 0) {
                            vscode.window.showWarningMessage('XLIDE Smoke: project has no VBA modules.');
                            return;
                        }

                        // Step 2: readModule (prefer a non-document module)
                        const target = modules.find(m => m.type !== 'document') ?? modules[0];
                        const source = await bridge.call<string>(
                            'readModule', { path: projectPath, module: target.name, full: false },
                        );
                        log(`[smoke] readModule "${target.name}" OK - ${source.length} chars`);

                        log('[smoke] All checks passed.');
                        void vscode.window.showInformationMessage(
                            `XLIDE Smoke: OK - ${modules.length} modules, read "${target.name}" (${source.length} chars). See XLIDE Output for details.`,
                        );
                    } catch (err) {
                        const msg = errorMessage(err);
                        log(`[smoke] FAILED: ${msg}`);
                        vscode.window.showErrorMessage(`XLIDE Smoke FAILED: ${msg}`);
                    }
                },
            );
        }),

        // Open the workbook in Excel (editable)
        registerXlideCommand('xlide.openWorkbook', async (node: XlideNode) => {
            const filePath = resolveProjectPath(node);
            if (!filePath) { return; }
            try {
                const attachToRunning = shouldAttachToRunningExcel();
                log(`[openWorkbook] Requested for: ${filePath}`);
                if (process.platform === 'win32') {
                    runWindowsExcel(filePath, attachToRunning, false);
                } else if (process.platform === 'darwin') {
                    cp.spawn('open', ['-a', 'Microsoft Excel', filePath])
                        .on('error', (err) => void vscode.window.showErrorMessage(`XLIDE: Could not open the project: ${errorMessage(err)}`));
                } else {
                    cp.spawn('libreoffice', ['--calc', '--norestore', filePath])
                        .on('error', (err) => void vscode.window.showErrorMessage(`XLIDE: Could not open the project: ${errorMessage(err)}`));
                }
            } catch (err) {
                vscode.window.showErrorMessage(`Failed to open project: ${err}`);
            }
        }),

        // Open the workbook in Excel (read-only)
        registerXlideCommand('xlide.openWorkbookReadOnly', async (node: XlideNode) => {
            const filePath = resolveProjectPath(node);
            if (!filePath) { return; }
            try {
                const attachToRunning = shouldAttachToRunningExcel();
                log(`[openWorkbookReadOnly] Requested for: ${filePath}`);
                if (process.platform === 'win32') {
                    runWindowsExcel(filePath, attachToRunning, true);
                } else if (process.platform === 'darwin') {
                    cp.spawn('open', ['-a', 'Microsoft Excel', filePath])
                        .on('error', (err) => void vscode.window.showErrorMessage(`XLIDE: Could not open the project: ${errorMessage(err)}`));
                } else {
                    cp.spawn('libreoffice', ['--calc', '--norestore', '--view', filePath])
                        .on('error', (err) => void vscode.window.showErrorMessage(`XLIDE: Could not open the project: ${errorMessage(err)}`));
                }
            } catch (err) {
                vscode.window.showErrorMessage(`Failed to open project: ${err}`);
            }
        }),

        // Open a non-Excel macro container in whatever application owns it
        // (Word, PowerPoint, Access): the OS association is the router.
        registerXlideCommand('xlide.openInOfficeApp', async (node: XlideNode) => {
            const filePath = resolveProjectPath(node);
            if (!filePath) { return; }
            log(`[openInOfficeApp] Requested for: ${filePath}`);
            const opened = await vscode.env.openExternal(vscode.Uri.file(filePath));
            if (!opened) {
                vscode.window.showErrorMessage(`XLIDE: Could not open ${path.basename(filePath)} in its Office application.`);
            }
        }),

        // Detect the Sub/Function at the cursor and open the project, then guide to run it
        registerXlideCommand('xlide.runMacroAtCursor', async () => {
            const editor = vscode.window.activeTextEditor;
            if (!editor || editor.document.uri.scheme !== XLIDE_SCHEME) {
                vscode.window.showWarningMessage('XLIDE: Open a VBA module to run a macro.');
                return;
            }

            try {
                // Decode the URI to get filePath and moduleName
                const { projectPath, moduleName } = decodeModuleUri(editor.document.uri);
                log(`[runMacro] Requested from module: ${moduleName} in ${projectPath}`);

                // Find which procedure the cursor is in (parser-based, so
                // Friend/Global/Static modifiers are recognized too). Done before
                // saving so a no-op cursor position bails out without a save.
                const procedure = procedureAtCursor(editor);
                if (!procedure) {
                    vscode.window.showWarningMessage('XLIDE: Cursor is not inside a Sub or Function.');
                    return;
                }
                // F5 runs the macro with no arguments (Application.Run with no
                // args), so a procedure with required parameters cannot run - VBA
                // rejects the call. Refuse up front with a clear message instead of
                // surfacing the opaque COM failure.
                const required = requiredParameterNames(procedure);
                if (required.length > 0) {
                    vscode.window.showWarningMessage(
                        `XLIDE: "${procedure.name}" has required parameter${required.length > 1 ? 's' : ''} `
                        + `(${required.join(', ')}) and cannot be run with F5, which passes no arguments. `
                        + `Make ${required.length > 1 ? 'them' : 'it'} Optional, or call it from a parameterless Sub.`,
                    );
                    return;
                }
                const currentProc = procedure.name;

                // The run machinery below is Excel COM end to end (launcher,
                // coordinator, reopen tracking), so gate by the file's host
                // first. Word and PowerPoint modules save and open in their
                // own application with run guidance; Access cannot run edits
                // at all, for the stated engine reason.
                const containerHost = containerHostForPath(projectPath);
                if (containerHost === 'access') {
                    vscode.window.showWarningMessage(
                        'XLIDE: Access files are read-only in XLIDE because Access runs compiled '
                        + 'p-code, so XLIDE cannot run this macro. Open the database in Access to run it.',
                    );
                    return;
                }
                if (containerHost === 'word' || containerHost === 'powerpoint') {
                    const app = containerAppNameForPath(projectPath);
                    if (editor.document.isDirty) {
                        await editor.document.save();
                    }
                    if (process.platform === 'win32') {
                        // Full parity with the Excel path: reopen read-only in
                        // the visible owning application and run the macro
                        // through its COM, per the harness-measured semantics.
                        try {
                            await runHostFileMacroReadOnly(
                                containerHost, projectPath, `${moduleName}.${currentProc}`, log,
                            );
                        } catch (err) {
                            showRunMacroFailure(err, app);
                        }
                        return;
                    }
                    // No COM off Windows: open in the owning application with
                    // guidance naming the exact macro.
                    const opened = await vscode.env.openExternal(vscode.Uri.file(projectPath));
                    if (!opened) {
                        vscode.window.showErrorMessage(
                            `XLIDE: Could not open ${path.basename(projectPath)} in ${app}.`,
                        );
                        return;
                    }
                    vscode.window.showInformationMessage(
                        `XLIDE: Opened in ${app}. Run "${moduleName}.${currentProc}" with Alt+F8 or from the VBE.`,
                    );
                    return;
                }

                // Suppress XLIDE's own post-save reopen for THIS workbook across the
                // whole run: F5 saves the dirty module and then reopens it read-only
                // itself to run the macro. Holding suppression over the save AND the
                // macro run keeps any refresh (this save's, or a concurrent save of
                // another module in the same workbook) from racing that reopen.
                await withWorkbookReopenSuppressed(projectPath, async () => {
                    // Persist in-editor changes first so the macro reflects the
                    // current source rather than the last-saved version.
                    if (editor.document.isDirty) {
                        await editor.document.save();
                    }

                    // Open the project read-only
                    if (process.platform === 'win32') {
                        const attachToRunning = shouldAttachToRunningExcel();
                        const macroRef = `${moduleName}.${currentProc}`;
                        log(`[runMacro] attachToRunningExcel=${attachToRunning}`);
                        try {
                            await runWorkbookMacroReadOnly(projectPath, macroRef, { attachToRunning }, log);
                        } catch (err) {
                            // The workbook is open for editing in Excel (locked). Honor
                            // the coordination policy: close it and retry (the macro
                            // script then reopens read-only to run) instead of asking
                            // the user to close it by hand. block mode still rethrows.
                            const settings = resolveExcelCoordinationSettings();
                            if (err instanceof ExcelMacroError && err.code === 'REOPEN_BLOCKED'
                                && settings.mode !== 'block' && shouldAttemptClose(settings, projectPath)) {
                                log(`[runMacro] reopen blocked; coordinationMode=${settings.mode}, closing workbook`);
                                await closeWorkbookInExcel(projectPath, { force: settings.mode === 'closeForce' }, log);
                                // The macro host is about to reopen the workbook read-only on
                                // retry; record it now so it stays tracked even if the macro
                                // itself then errors (RUN_FAILED) before we mark below.
                                markWorkbookOpenedByXlide(projectPath);
                                await runWorkbookMacroReadOnly(projectPath, macroRef, { attachToRunning }, log);
                            } else {
                                // RUN_FAILED means the macro host already reopened the
                                // project read-only before the macro raised, so record it
                                // (mirroring the post-success mark below) so a later
                                // closeTracked save still frees the lock. Then rethrow.
                                if (err instanceof ExcelMacroError && err.code === 'RUN_FAILED') {
                                    markWorkbookOpenedByXlide(projectPath);
                                }
                                throw err;
                            }
                        }
                        // The macro host reopened the workbook read-only; record it so
                        // a later closeTracked save can free the lock automatically.
                        markWorkbookOpenedByXlide(projectPath);
                    } else if (process.platform === 'darwin') {
                        cp.spawn('open', ['-a', 'Microsoft Excel', projectPath])
                            .on('error', (err) => void vscode.window.showErrorMessage(`XLIDE: Could not open the project: ${errorMessage(err)}`));
                        vscode.window.showInformationMessage(
                            `Workbook opened. Run macro: ${moduleName}.${currentProc}`,
                        );
                    } else {
                        cp.spawn('libreoffice', ['--calc', '--norestore', '--view', projectPath])
                            .on('error', (err) => void vscode.window.showErrorMessage(`XLIDE: Could not open the project: ${errorMessage(err)}`));
                        vscode.window.showInformationMessage(
                            `Workbook opened. Run macro manually: ${moduleName}.${currentProc}`,
                        );
                    }
                });
            } catch (err) {
                showRunMacroFailure(err);
            }
        }),
    ];
}
