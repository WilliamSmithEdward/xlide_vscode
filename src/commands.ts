import * as vscode from 'vscode';
import * as path from 'path';
import * as cp from 'child_process';
import { PythonBridge } from './pythonBridge';
import { XlsmExplorer, XlideNode } from './xlsmExplorer';
import { XlideFileSystemProvider, encodeModuleUri, decodeModuleUri, XLIDE_SCHEME } from './xlideFileSystem';
import { encodeRemoteModuleUri } from './liveShare';
import {
    type ExportMode,
    exportWorkbookModules,
    normalizeExportMode,
    readWorkbookRepoConfig,
    writeWorkbookRepoConfig,
    setWorkbookExportMode,
} from './moduleDump';

function psSingleQuoted(value: string): string {
    return `'${value.replace(/'/g, "''")}'`;
}

export function registerCommands(
    _context: vscode.ExtensionContext,
    bridge: PythonBridge,
    explorer: XlsmExplorer,
    _fsProvider: XlideFileSystemProvider,
    out: vscode.OutputChannel,
): vscode.Disposable[] {
    function log(msg: string): void {
        out.appendLine(msg);
        out.show(true);
    }

    function shouldAttachToRunningExcel(): boolean {
        return vscode.workspace
            .getConfiguration('xlide')
            .get<boolean>('attachToRunningExcel', true);
    }

    // Helper functions for Windows COM-based Excel operations
    function runWindowsExcelReadOnly(filePath: string, attachToRunning: boolean): void {
        const script = [
            '$ErrorActionPreference = "Stop"',
            `$targetPath = ${psSingleQuoted(filePath)}`,
            `$targetName = ${psSingleQuoted(path.basename(filePath))}`,
            '$excel = $null',
            '$workbook = $null',
            `$attachToRunning = ${attachToRunning ? '$true' : '$false'}`,
            'if ($attachToRunning) {',
            '  try { $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application") } catch { }',
            '}',
            'if (-not $excel) {',
            '  $excel = New-Object -ComObject Excel.Application',
            '}',
            '$excel.Visible = $true',
            'foreach ($wb in @($excel.Workbooks)) {',
            '  if (($wb.FullName -ieq $targetPath) -or ($wb.Name -ieq $targetName)) { $workbook = $wb; break }',
            '}',
            'if (-not $workbook) {',
            '  $workbook = $excel.Workbooks.Open($targetPath, 0, $true)',
            '}',
            '$workbook.Activate()',
        ].join('; ');

        log(`[openWorkbook] Running: powershell -Command "${script}"`);
        const child = cp.spawn('powershell.exe', [
            '-NoProfile',
            '-ExecutionPolicy',
            'Bypass',
            '-Command',
            script,
        ]);
        child.on('spawn', () => {
            log(`[openWorkbook] Spawned powershell.exe (pid=${child.pid ?? 'unknown'})`);
        });
        child.on('error', (err) => {
            log(`[openWorkbook] Error: ${err.message}`);
            void vscode.window.showErrorMessage(`XLIDE: Open Workbook failed: ${err.message}`);
        });
        child.stdout?.on('data', (d: Buffer) => {
            const text = d.toString().trim();
            if (text) {
                log(`[openWorkbook stdout] ${text}`);
            }
        });
        child.stderr?.on('data', (d: Buffer) => {
            const text = d.toString().trim();
            if (text) {
                log(`[openWorkbook stderr] ${text}`);
            }
        });
        child.on('exit', (code, signal) => {
            log(`[openWorkbook] powershell exited with code=${code} signal=${signal ?? 'none'}`);
        });
    }

    function runWindowsExcelMacroReadOnly(filePath: string, macroName: string, attachToRunning: boolean): void {
        const script = [
            '$ErrorActionPreference = "Stop"',
            `$targetPath = ${psSingleQuoted(filePath)}`,
            `$targetName = ${psSingleQuoted(path.basename(filePath))}`,
            `$macroName = ${psSingleQuoted(macroName)}`,
            '$excel = $null',
            '$workbook = $null',
            `$attachToRunning = ${attachToRunning ? '$true' : '$false'}`,
            'if ($attachToRunning) {',
            '  try { $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application") } catch { }',
            '}',
            'if (-not $excel) {',
            '  $excel = New-Object -ComObject Excel.Application',
            '}',
            '$excel.Visible = $true',
            'foreach ($wb in @($excel.Workbooks)) {',
            '  if (($wb.FullName -ieq $targetPath) -or ($wb.Name -ieq $targetName)) { $workbook = $wb; break }',
            '}',
            'if (-not $workbook) {',
            '  $workbook = $excel.Workbooks.Open($targetPath, 0, $true)',
            '}',
            '$workbook.Activate()',
            '$macroRef = "\'" + $workbook.Name + "\'!" + $macroName',
            '$excel.Run($macroRef)',
        ].join('; ');

        log(`[runMacro] Running: ${macroName}`);
        log(`[runMacro] Script: ${script}`);
        const child = cp.spawn('powershell.exe', [
            '-NoProfile',
            '-ExecutionPolicy',
            'Bypass',
            '-Command',
            script,
        ]);
        child.on('spawn', () => {
            log(`[runMacro] Spawned powershell.exe (pid=${child.pid ?? 'unknown'})`);
        });
        child.on('error', (err) => {
            log(`[runMacro] Error: ${err.message}`);
            void vscode.window.showErrorMessage(`XLIDE: Run Macro failed: ${err.message}`);
        });
        child.stdout?.on('data', (d: Buffer) => {
            const text = d.toString().trim();
            if (text) {
                log(`[runMacro stdout] ${text}`);
            }
        });
        child.stderr?.on('data', (d: Buffer) => {
            const text = d.toString().trim();
            if (text) {
                log(`[runMacro stderr] ${text}`);
            }
        });
        child.on('exit', (code, signal) => {
            log(`[runMacro] powershell exited with code=${code} signal=${signal ?? 'none'}`);
        });
    }

    function resolveWorkbookPath(node?: XlideNode): string | undefined {
        let filePath = node?.filePath;
        if (!filePath) {
            const active = vscode.window.activeTextEditor;
            if (active && active.document.uri.scheme === XLIDE_SCHEME) {
                filePath = decodeModuleUri(active.document.uri).xlsmPath;
            }
        }
        return filePath;
    }

    return [
        vscode.commands.registerCommand('xlide.refreshExplorer', () => {
            explorer.refresh();
        }),

        // Open a module (or navigate to a sub's line inside one)
        vscode.commands.registerCommand('xlide.openModule', async (node: XlideNode) => {
            if (!node?.moduleName) { return; }
            const uri = node.isRemote && node.remoteId
                ? encodeRemoteModuleUri(node.remoteId, node.moduleName)
                : encodeModuleUri(node.filePath, node.moduleName);

            // Set the language to 'vba' so syntax highlighters kick in
            const doc = await vscode.workspace.openTextDocument(uri);
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

            // Set language mode to vba for the document
            await vscode.languages.setTextDocumentLanguage(doc, 'vba');
        }),

        // Add a new standard module to an .xlsm
        vscode.commands.registerCommand('xlide.newModule', async (node: XlideNode) => {
            if (node?.kind !== 'xlsm') { return; }
            const name = await vscode.window.showInputBox({
                prompt: 'New module name',
                placeHolder: 'Module1',
                validateInput: (v) =>
                    /^\w+$/.test(v) ? undefined : 'Module names must be alphanumeric',
            });
            if (!name) { return; }

            const stub = `Option Explicit\r\n\r\nSub ${name}_Main()\r\n\r\nEnd Sub\r\n`;
            try {
                await bridge.call('writeModule', {
                    path: node.filePath,
                    module: name,
                    source: stub,
                });
                explorer.refresh();
                // Open the new module immediately
                const uri = encodeModuleUri(node.filePath, name);
                const doc = await vscode.workspace.openTextDocument(uri);
                await vscode.window.showTextDocument(doc, { preview: false });
                await vscode.languages.setTextDocumentLanguage(doc, 'vba');
            } catch (err) {
                vscode.window.showErrorMessage(`XLIDE: Failed to create module: ${err}`);
            }
        }),

        // Rename a module
        vscode.commands.registerCommand('xlide.renameModule', async (node: XlideNode) => {
            if (!node?.moduleName) { return; }
            const newName = await vscode.window.showInputBox({
                prompt: `Rename "${node.moduleName}" to`,
                value: node.moduleName,
                validateInput: (v) =>
                    /^\w+$/.test(v) ? undefined : 'Module names must be alphanumeric',
            });
            if (!newName || newName === node.moduleName) { return; }

            try {
                await bridge.call('renameModule', {
                    path: node.filePath,
                    module: node.moduleName,
                    newName,
                });
                explorer.refresh();
            } catch (err) {
                vscode.window.showErrorMessage(`XLIDE: Rename failed: ${err}`);
            }
        }),

        // Delete a module (with confirmation)
        vscode.commands.registerCommand('xlide.deleteModule', async (node: XlideNode) => {
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
                await bridge.call('deleteModule', {
                    path: node.filePath,
                    module: node.moduleName,
                });
                // Close any open editors for this module
                const uri = encodeModuleUri(node.filePath, node.moduleName);
                for (const tab of vscode.window.tabGroups.all.flatMap((g) => g.tabs)) {
                    const input = tab.input;
                    if (
                        input instanceof vscode.TabInputText &&
                        input.uri.toString() === uri.toString()
                    ) {
                        await vscode.window.tabGroups.close(tab);
                    }
                }
                explorer.refresh();
            } catch (err) {
                vscode.window.showErrorMessage(`XLIDE: Delete failed: ${err}`);
            }
        }),

        // Export all modules to a user-selected folder and persist folder in workbook config JSON
        vscode.commands.registerCommand('xlide.exportModulesToFolder', async (node: XlideNode) => {
            const filePath = resolveWorkbookPath(node);
            if (!filePath) { return; }

            try {
                const existingConfig = await readWorkbookRepoConfig(filePath);
                const exportMode = normalizeExportMode(existingConfig.exportMode ?? existingConfig.dumpMode);
                const configuredFolder = existingConfig.exportFolder ?? existingConfig.dumpFolder;

                let exportFolder: string;
                if (configuredFolder) {
                    // Folder already set — export directly without prompting
                    exportFolder = configuredFolder;
                } else {
                    // First time — ask the user to pick a folder
                    const selected = await vscode.window.showOpenDialog({
                        canSelectFiles: false,
                        canSelectFolders: true,
                        canSelectMany: false,
                        openLabel: 'Select export folder',
                        defaultUri: vscode.Uri.file(path.dirname(filePath)),
                    });
                    if (!selected || selected.length === 0) { return; }
                    exportFolder = selected[0].fsPath;
                }

                log(`[exportModules] Workbook: ${filePath}`);
                log(`[exportModules] Target folder: ${exportFolder}`);
                log(`[exportModules] Mode: ${exportMode}`);

                const result = await exportWorkbookModules(bridge, {
                    filePath,
                    exportFolder,
                    exportMode,
                });

                log(`[exportModules] Wrote ${result.writtenCount} module(s)`);
                if (result.skippedNewCount > 0) {
                    log(`[exportModules] Skipped ${result.skippedNewCount} new module(s) because mode=replaceExistingOnly`);
                }
                if (result.removedCount > 0) {
                    log(`[exportModules] Removed ${result.removedCount} stale module file(s)`);
                }
                log(`[exportModules] Config updated: ${result.configPath}`);
                vscode.window.showInformationMessage(
                    `XLIDE: Exported ${result.writtenCount} module(s) to ${result.exportFolder} [mode=${result.exportMode}]`,
                );
            } catch (err) {
                const message = err instanceof Error ? err.message : String(err);
                log(`[exportModules] Error: ${message}`);
                vscode.window.showErrorMessage(`XLIDE: Failed to export modules: ${message}`);
            }
        }),

        // Backward-compatible alias for previous command id
        vscode.commands.registerCommand('xlide.dumpModulesToFolder', async (node: XlideNode) => {
            await vscode.commands.executeCommand('xlide.exportModulesToFolder', node);
        }),

        // Change the configured export folder for this workbook
        vscode.commands.registerCommand('xlide.changeRepoFolder', async (node: XlideNode) => {
            const filePath = resolveWorkbookPath(node);
            if (!filePath) { return; }

            try {
                const existingConfig = await readWorkbookRepoConfig(filePath);
                const currentFolder = existingConfig.exportFolder ?? existingConfig.dumpFolder;
                const selected = await vscode.window.showOpenDialog({
                    canSelectFiles: false,
                    canSelectFolders: true,
                    canSelectMany: false,
                    openLabel: 'Select new export folder',
                    defaultUri: currentFolder
                        ? vscode.Uri.file(currentFolder)
                        : vscode.Uri.file(path.dirname(filePath)),
                });
                if (!selected || selected.length === 0) { return; }

                const newFolder = selected[0].fsPath;
                await writeWorkbookRepoConfig(filePath, {
                    ...existingConfig,
                    exportFolder: newFolder,
                    dumpFolder: undefined,
                });
                log(`[changeRepoFolder] Folder set to ${newFolder} for ${filePath}`);
                vscode.window.showInformationMessage(
                    `XLIDE: Export folder updated to ${newFolder}`,
                );
            } catch (err) {
                const message = err instanceof Error ? err.message : String(err);
                vscode.window.showErrorMessage(`XLIDE: Failed to update export folder: ${message}`);
            }
        }),

        // Configure export behavior for this workbook
        vscode.commands.registerCommand('xlide.configureExportMode', async (node: XlideNode) => {
            const filePath = resolveWorkbookPath(node);
            if (!filePath) { return; }

            try {
                const existingConfig = await readWorkbookRepoConfig(filePath);
                const currentMode = normalizeExportMode(existingConfig.exportMode ?? existingConfig.dumpMode);
                const selection = await vscode.window.showQuickPick(
                    [
                        {
                            label: 'True Up (default)',
                            description: 'Replace existing, add new, remove no longer existing',
                            mode: 'trueUp' as ExportMode,
                        },
                        {
                            label: 'Replace Existing Only',
                            description: 'Replace files that already exist in the folder only',
                            mode: 'replaceExistingOnly' as ExportMode,
                        },
                    ],
                    {
                        title: `Configure module export mode for ${path.basename(filePath)}`,
                        placeHolder: currentMode === 'trueUp'
                            ? 'Current: True Up'
                            : 'Current: Replace Existing Only',
                    },
                );

                if (!selection) { return; }

                await setWorkbookExportMode(filePath, selection.mode);

                log(`[exportModules] Config mode set to ${selection.mode} for ${filePath}`);
                vscode.window.showInformationMessage(
                    `XLIDE: Export mode set to ${selection.mode} for ${path.basename(filePath)}`,
                );
            } catch (err) {
                const message = err instanceof Error ? err.message : String(err);
                log(`[exportModules] Configure mode error: ${message}`);
                vscode.window.showErrorMessage(`XLIDE: Failed to configure export mode: ${message}`);
            }
        }),

        // Backward-compatible alias for previous command id
        vscode.commands.registerCommand('xlide.configureDumpMode', async (node: XlideNode) => {
            await vscode.commands.executeCommand('xlide.configureExportMode', node);
        }),

        // Open the workbook with the registered app (Excel), read-only
        vscode.commands.registerCommand('xlide.openWorkbook', async (node: XlideNode) => {
            const filePath = resolveWorkbookPath(node);
            if (!filePath) { return; }

            try {
                const attachToRunning = shouldAttachToRunningExcel();
                log(`[openWorkbook] Requested for: ${filePath}`);
                log(`[openWorkbook] attachToRunningExcel=${attachToRunning}`);
                if (process.platform === 'win32') {
                    runWindowsExcelReadOnly(filePath, attachToRunning);
                } else if (process.platform === 'darwin') {
                    // macOS: use open with Excel (read-only requires AppleScript, so just open normally)
                    cp.spawn('open', ['-a', 'Microsoft Excel', filePath]);
                } else {
                    // Linux: libreoffice with read-only flag
                    cp.spawn('libreoffice', ['--calc', '--norestore', '--view', filePath]);
                }
            } catch (err) {
                vscode.window.showErrorMessage(`Failed to open workbook: ${err}`);
            }
        }),

        // Detect the Sub/Function at the cursor and open the workbook, then guide to run it
        vscode.commands.registerCommand('xlide.runMacroAtCursor', async () => {
            const editor = vscode.window.activeTextEditor;
            if (!editor || !editor.document.uri.scheme.startsWith(XLIDE_SCHEME)) {
                vscode.window.showWarningMessage('XLIDE: Open a VBA module to run a macro.');
                return;
            }

            try {
                // Decode the URI to get filePath and moduleName
                const { xlsmPath, moduleName } = decodeModuleUri(editor.document.uri);
                log(`[runMacro] Requested from module: ${moduleName} in ${xlsmPath}`);

                // Get the source code and find which Sub/Function the cursor is in
                const result = await bridge.call<{ source: string }>(
                    'readModule',
                    { path: xlsmPath, module: moduleName },
                );

                const cursorLine = editor.selection.active.line;
                const source = result.source;
                const lines = source.split('\n');

                // Find the current Sub/Function
                const procRe = /^\s*(Public|Private)?\s*(Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)/i;
                let currentProc = '';
                for (let i = cursorLine; i >= 0; i--) {
                    const match = lines[i].match(procRe);
                    if (match) {
                        currentProc = match[3];
                        break;
                    }
                }

                if (!currentProc) {
                    vscode.window.showWarningMessage('XLIDE: Cursor is not inside a Sub or Function.');
                    return;
                }

                // Open the workbook read-only
                if (process.platform === 'win32') {
                    const attachToRunning = shouldAttachToRunningExcel();
                    log(`[runMacro] attachToRunningExcel=${attachToRunning}`);
                    runWindowsExcelMacroReadOnly(xlsmPath, `${moduleName}.${currentProc}`, attachToRunning);
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
                vscode.window.showErrorMessage(`XLIDE: Failed to run macro: ${err}`);
            }
        }),
    ];
}
