// Commands that launch the application owning the file - Excel, Word,
// PowerPoint or Access - and run a macro in it.
//
// Desktop only, by construction: there is no Office to launch from a
// browser. platformFeaturesNode registers these and platformFeaturesWeb
// does not, which is also what keeps child_process out of the web bundle.

import * as vscode from 'vscode';
import * as path from 'path';
import * as cp from 'child_process';
import {
    decodeModuleUri,
    XLIDE_SCHEME,
} from '../xlideFileSystem';
import { xlideOfficeAttachToRunningFromConfig } from '../globalSettings';
import { containerAppNameForPath } from '../macroContainerUi';
import { OFFICE_HOST_APPS, officeHostForPath } from '../officeHostApps';
import { registerXlideCommand } from '../xlideCommandRegistration';
import type { XlideNode } from '../projectExplorer';
import { errorMessage } from '../util/errors';
import {
    HostMacroError,
    openFileInHost,
    runHostMacro,
} from '../officeHostLauncher';
import {
    closeFileInHost,
    markFileOpenedByXlide,
    resolveHostCoordinationSettings,
    shouldAttemptClose,
    withFileReopenSuppressed,
} from '../officeWriteCoordinator';
import {
    procedureAtCursor,
    requiredParameterNames,
    resolveProjectPath,
    type CommandDeps,
    outputLogger,
} from './shared';

export function registerOfficeAppCommands(deps: CommandDeps): vscode.Disposable[] {
    const { out } = deps;

    const log = outputLogger(out);

    /** Hands the file to whatever the operating system has registered for it. */
    async function openThroughShell(filePath: string): Promise<boolean> {
        const opened = await vscode.env.openExternal(vscode.Uri.file(filePath));
        if (!opened) {
            vscode.window.showErrorMessage(
                `XLIDE: Could not open ${path.basename(filePath)} in ${containerAppNameForPath(filePath)}.`,
            );
        }
        return opened;
    }

    /** Opens the file in the application that owns it, read-only when asked. */
    const openCommand = (readOnly: boolean, tag: string) => async (node: XlideNode): Promise<void> => {
        const filePath = resolveProjectPath(node);
        if (filePath) {
            openInHostApp(filePath, readOnly, tag);
        }
    };

    function shouldAttachToRunningApp(): boolean {
        return xlideOfficeAttachToRunningFromConfig(vscode.workspace.getConfiguration('xlide')).value;
    }

    /**
     * Starts the file's application off Windows, where there is no COM: the
     * application bundle on macOS, LibreOffice's matching module on Linux.
     * Answers false when the platform has neither for this host.
     */
    function spawnHostApp(filePath: string, readOnly: boolean): boolean {
        const host = officeHostForPath(filePath);
        const info = host ? OFFICE_HOST_APPS[host] : undefined;
        const onError = (err: Error): void => void vscode.window.showErrorMessage(
            `XLIDE: Could not open the project: ${errorMessage(err)}`,
        );
        if (process.platform === 'darwin' && info?.macAppName) {
            cp.spawn('open', ['-a', info.macAppName, filePath]).on('error', onError);
            return true;
        }
        if (process.platform !== 'darwin' && info?.libreOfficeFlag) {
            cp.spawn('libreoffice', [info.libreOfficeFlag, '--norestore', ...(readOnly ? ['--view'] : []), filePath])
                .on('error', onError);
            return true;
        }
        return false;
    }

    function openInHostApp(filePath: string, readOnly: boolean, tag: string): void {
        log(`[${tag}] Requested for: ${filePath}`);
        try {
            if (process.platform === 'win32' && officeHostForPath(filePath)) {
                // Remember XLIDE opened this file so closeTracked coordination can
                // later close it without touching files the user opened manually.
                markFileOpenedByXlide(filePath);
                const app = containerAppNameForPath(filePath);
                void openFileInHost(filePath, { attachToRunning: shouldAttachToRunningApp(), readOnly }, log)
                    .catch((err: Error) => {
                        void vscode.window.showErrorMessage(
                            `XLIDE: Could not open ${path.basename(filePath)} in ${app}: ${err.message}`,
                        );
                    });
                return;
            }
            if (process.platform === 'win32' || !spawnHostApp(filePath, readOnly)) {
                void openThroughShell(filePath);
            }
        } catch (err) {
            vscode.window.showErrorMessage(`Failed to open project: ${err}`);
        }
    }

    function showRunMacroFailure(err: unknown, appName: string): void {
        if (err instanceof HostMacroError &&
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

    return [
            // Open the file in the application that owns it: Excel, Word,
            // PowerPoint or Access.
            registerXlideCommand('xlide.openInOfficeApp', openCommand(false, 'openInOfficeApp')),
            registerXlideCommand('xlide.openInOfficeAppReadOnly', openCommand(true, 'openInOfficeAppReadOnly')),
            // The ids these two commands had while they only opened Excel. They
            // stay registered, undeclared, so a keybinding made against them
            // keeps working.
            registerXlideCommand('xlide.openWorkbook', openCommand(false, 'openInOfficeApp')),
            registerXlideCommand('xlide.openWorkbookReadOnly', openCommand(true, 'openInOfficeAppReadOnly')),

            // Detect the Sub/Function at the cursor and open the project, then guide to run it
            registerXlideCommand('xlide.runMacroAtCursor', async () => {
                const editor = vscode.window.activeTextEditor;
                if (!editor || editor.document.uri.scheme !== XLIDE_SCHEME) {
                    vscode.window.showWarningMessage('XLIDE: Open a VBA module to run a macro.');
                    return;
                }

                // Named once the file is known, for the failure message below.
                let appName = 'the application';
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

                    const app = containerAppNameForPath(projectPath);
                    appName = app;
                    const macro = { moduleName, procedureName: currentProc };

                    // Suppress XLIDE's own post-save reopen for THIS file across the
                    // whole run: F5 saves the dirty module and then reopens the file
                    // itself to run the macro. Holding suppression over the save AND the
                    // macro run keeps any reopen (this save's, or a concurrent save of
                    // another module in the same file) from racing that one.
                    await withFileReopenSuppressed(projectPath, async () => {
                        // Persist in-editor changes first so the macro reflects the
                        // current source rather than the last-saved version. A save
                        // that failed has already said why, and running on would run
                        // the stale code.
                        if (editor.document.isDirty && !(await editor.document.save())) {
                            log('[runMacro] the module could not be saved; not running');
                            return;
                        }

                        if (process.platform === 'win32' && officeHostForPath(projectPath)) {
                            // Open in the visible owning application and run the macro
                            // through its COM, per the harness-measured semantics of
                            // each host (officeHostLauncher.ts).
                            const attachToRunning = shouldAttachToRunningApp();
                            log(`[runMacro] attachToRunning=${attachToRunning}`);
                            try {
                                await runHostMacro(projectPath, macro, { attachToRunning }, log);
                            } catch (err) {
                                // The file is open for editing in its application (locked).
                                // Honor the coordination policy: close it and retry (the
                                // macro script then reopens it to run) instead of asking
                                // the user to close it by hand. block mode still rethrows.
                                const settings = resolveHostCoordinationSettings();
                                if (err instanceof HostMacroError && err.code === 'REOPEN_BLOCKED'
                                    && settings.mode !== 'block' && shouldAttemptClose(settings, projectPath)) {
                                    log(`[runMacro] reopen blocked; coordinationMode=${settings.mode}, closing in ${app}`);
                                    await closeFileInHost(projectPath, { force: settings.mode === 'closeForce' }, log);
                                    // The macro host is about to reopen the file on retry;
                                    // record it now so it stays tracked even if the macro
                                    // itself then errors (RUN_FAILED) before we mark below.
                                    markFileOpenedByXlide(projectPath);
                                    await runHostMacro(projectPath, macro, { attachToRunning }, log);
                                } else {
                                    // RUN_FAILED means the macro host already reopened the
                                    // project before the macro raised, so record it
                                    // (mirroring the post-success mark below) so a later
                                    // save can still free the lock. Then rethrow.
                                    if (err instanceof HostMacroError && err.code === 'RUN_FAILED') {
                                        markFileOpenedByXlide(projectPath);
                                    }
                                    throw err;
                                }
                            }
                            // The macro host left the file open; record it so a later
                            // save can free the lock automatically.
                            markFileOpenedByXlide(projectPath);
                            return;
                        }

                        // No COM off Windows: open in the owning application, read-only
                        // where the platform can, with guidance naming the exact macro.
                        if (!spawnHostApp(projectPath, true) && !(await openThroughShell(projectPath))) {
                            return;
                        }
                        vscode.window.showInformationMessage(
                            `XLIDE: Opened ${path.basename(projectPath)}. Run "${moduleName}.${currentProc}" from the application's macro dialog.`,
                        );
                    });
                } catch (err) {
                    showRunMacroFailure(err, appName);
                }
            }),
        ];
    }
