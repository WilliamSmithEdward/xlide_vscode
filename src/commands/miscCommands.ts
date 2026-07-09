import * as vscode from 'vscode';
import * as path from 'path';
import * as cp from 'child_process';
import {
    encodeModuleUri,
    decodeModuleUri,
    XLIDE_SCHEME,
    XLIDE_VBA_LANGUAGE_ID,
} from '../xlideFileSystem';
import { encodeRemoteModuleUri } from '../liveShare';
import { xlideAttachToRunningExcelFromConfig } from '../globalSettings';
import { registerXlideCommand } from '../xlideCommandRegistration';
import type { XlideNode } from '../xlsmExplorer';
import { errorMessage } from '../util/errors';
import {
    ExcelMacroError,
    openWorkbookInExcel,
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
    resolveWorkbookPath,
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

    function showRunMacroFailure(err: unknown): void {
        if (err instanceof ExcelMacroError &&
            (err.code === 'REOPEN_BLOCKED' || err.code === 'REOPEN_FAILED')) {
            void vscode.window.showWarningMessage(`XLIDE: ${err.message}`);
            return;
        }
        const message = errorMessage(err);
        // Excel was busy and kept rejecting the COM call even after XLIDE's retries
        // (RPC_E_CALL_REJECTED / RETRYLATER) - almost always a modal dialog left
        // open in Excel, such as a MsgBox from a previous run.
        if (/rejected by callee|RPC_E_CALL_REJECTED|0x80010001|RETRYLATER|0x8001010A/i.test(message)) {
            void vscode.window.showWarningMessage(
                'XLIDE: Excel is busy, so the macro could not run. A dialog may be open in Excel '
                + '(for example a MsgBox from a previous run); close it, then press F5 again.',
            );
            return;
        }
        void vscode.window.showErrorMessage(`XLIDE: Failed to run macro: ${message}`);
    }

    // Windows COM-based Excel launch; the script lives in excelLauncher.ts.
    function runWindowsExcel(filePath: string, attachToRunning: boolean, readOnly: boolean): void {
        // Remember XLIDE opened this workbook so closeTracked coordination can
        // later close it without touching workbooks the user opened manually.
        markWorkbookOpenedByXlide(filePath);
        void openWorkbookInExcel(filePath, { attachToRunning, readOnly }, log).catch((err: Error) => {
            void vscode.window.showErrorMessage(`XLIDE: Open Workbook failed: ${err.message}`);
        });
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

        // In-tree "Load failed - click to retry" placeholder (e.g. after Excel
        // briefly held the workbook file). Retries just the failed listing
        // instead of collapsing the whole tree with a full refresh.
        registerXlideCommand('xlide.retryExplorerLoad', (node: XlideNode) => {
            explorer.retryLoad(node);
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
            await vscode.commands.executeCommand('references-view.findReferences', uri, pos);
        }),

        // DEV: smoke test - verifies listModules + readModule against a workspace workbook
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
                        log(`[smoke] listModules OK - ${modules.length} module(s): ${modules.map(m => m.name).join(', ')}`);

                        if (modules.length === 0) {
                            vscode.window.showWarningMessage('XLIDE Smoke: workbook has no VBA modules.');
                            return;
                        }

                        // Step 2: readModule (prefer a non-document module)
                        const target = modules.find(m => m.type !== 'document') ?? modules[0];
                        const source = await bridge.call<string>(
                            'readModule', { path: workbookPath, module: target.name, full: false },
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
            const filePath = resolveWorkbookPath(node);
            if (!filePath) { return; }
            try {
                const attachToRunning = shouldAttachToRunningExcel();
                log(`[openWorkbook] Requested for: ${filePath}`);
                if (process.platform === 'win32') {
                    runWindowsExcel(filePath, attachToRunning, false);
                } else if (process.platform === 'darwin') {
                    cp.spawn('open', ['-a', 'Microsoft Excel', filePath])
                        .on('error', (err) => void vscode.window.showErrorMessage(`XLIDE: Could not open the workbook: ${errorMessage(err)}`));
                } else {
                    cp.spawn('libreoffice', ['--calc', '--norestore', filePath])
                        .on('error', (err) => void vscode.window.showErrorMessage(`XLIDE: Could not open the workbook: ${errorMessage(err)}`));
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
                    cp.spawn('open', ['-a', 'Microsoft Excel', filePath])
                        .on('error', (err) => void vscode.window.showErrorMessage(`XLIDE: Could not open the workbook: ${errorMessage(err)}`));
                } else {
                    cp.spawn('libreoffice', ['--calc', '--norestore', '--view', filePath])
                        .on('error', (err) => void vscode.window.showErrorMessage(`XLIDE: Could not open the workbook: ${errorMessage(err)}`));
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
                // Decode the URI to get filePath and moduleName
                const { xlsmPath, moduleName } = decodeModuleUri(editor.document.uri);
                log(`[runMacro] Requested from module: ${moduleName} in ${xlsmPath}`);

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

                // Suppress XLIDE's own post-save reopen for THIS workbook across the
                // whole run: F5 saves the dirty module and then reopens it read-only
                // itself to run the macro. Holding suppression over the save AND the
                // macro run keeps any refresh (this save's, or a concurrent save of
                // another module in the same workbook) from racing that reopen.
                await withWorkbookReopenSuppressed(xlsmPath, async () => {
                    // Persist in-editor changes first so the macro reflects the
                    // current source rather than the last-saved version.
                    if (editor.document.isDirty) {
                        await editor.document.save();
                    }

                    // Open the workbook read-only
                    if (process.platform === 'win32') {
                        const attachToRunning = shouldAttachToRunningExcel();
                        const macroRef = `${moduleName}.${currentProc}`;
                        log(`[runMacro] attachToRunningExcel=${attachToRunning}`);
                        try {
                            await runWorkbookMacroReadOnly(xlsmPath, macroRef, { attachToRunning }, log);
                        } catch (err) {
                            // The workbook is open for editing in Excel (locked). Honor
                            // the coordination policy: close it and retry (the macro
                            // script then reopens read-only to run) instead of asking
                            // the user to close it by hand. block mode still rethrows.
                            const settings = resolveExcelCoordinationSettings();
                            if (err instanceof ExcelMacroError && err.code === 'REOPEN_BLOCKED'
                                && settings.mode !== 'block' && shouldAttemptClose(settings, xlsmPath)) {
                                log(`[runMacro] reopen blocked; coordinationMode=${settings.mode}, closing workbook`);
                                await closeWorkbookInExcel(xlsmPath, { force: settings.mode === 'closeForce' }, log);
                                // The macro host is about to reopen the workbook read-only on
                                // retry; record it now so it stays tracked even if the macro
                                // itself then errors (RUN_FAILED) before we mark below.
                                markWorkbookOpenedByXlide(xlsmPath);
                                await runWorkbookMacroReadOnly(xlsmPath, macroRef, { attachToRunning }, log);
                            } else {
                                // RUN_FAILED means the macro host already reopened the
                                // workbook read-only before the macro raised, so record it
                                // (mirroring the post-success mark below) so a later
                                // closeTracked save still frees the lock. Then rethrow.
                                if (err instanceof ExcelMacroError && err.code === 'RUN_FAILED') {
                                    markWorkbookOpenedByXlide(xlsmPath);
                                }
                                throw err;
                            }
                        }
                        // The macro host reopened the workbook read-only; record it so
                        // a later closeTracked save can free the lock automatically.
                        markWorkbookOpenedByXlide(xlsmPath);
                    } else if (process.platform === 'darwin') {
                        cp.spawn('open', ['-a', 'Microsoft Excel', xlsmPath])
                            .on('error', (err) => void vscode.window.showErrorMessage(`XLIDE: Could not open the workbook: ${errorMessage(err)}`));
                        vscode.window.showInformationMessage(
                            `Workbook opened. Run macro: ${moduleName}.${currentProc}`,
                        );
                    } else {
                        cp.spawn('libreoffice', ['--calc', '--norestore', '--view', xlsmPath])
                            .on('error', (err) => void vscode.window.showErrorMessage(`XLIDE: Could not open the workbook: ${errorMessage(err)}`));
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
