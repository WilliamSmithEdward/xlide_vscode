import * as vscode from 'vscode';
import * as path from 'path';
import * as cp from 'child_process';
import * as fs from 'fs';
import { PythonBridge } from './pythonBridge';
import { XlsmExplorer, XlideNode } from './xlsmExplorer';
import { XlideFileSystemProvider, encodeModuleUri, decodeModuleUri, XLIDE_SCHEME, notifySignatureDropped } from './xlideFileSystem';
import { encodeRemoteModuleUri } from './liveShare';
import {
    type ExportMode,
    exportWorkbookModules,
    normalizeExportMode,
    readWorkbookRepoConfig,
    writeWorkbookRepoConfig,
    setWorkbookExportMode,
} from './moduleExport';
import {
    lintWorkbook,
    summarizeWorkbookLintProblems,
    workbookProblemsForModule,
    type WorkbookLintProblem,
    type WorkbookLintSummary,
} from './vbaWorkbookLint';
import { lintVbaModuleSource } from './vbaModuleLint';
import { VBA_IDENTIFIER_NAME_RE } from './vbaLinter';
import { VbaSymbolIndex } from './vbaSymbolIndex';
import {
    buildVbaProjectIndex,
    moduleKindFromType,
    projectClassModuleDefinition,
} from './vbaNavigation';
import {
    projectAnalysisOptionsForModule,
    projectProcedureSignatures,
} from './vbaProjectAnalysis';
import {
    projectClassReferenceEdit,
    renameProjectClassModule,
} from './vbaClassRename';
import type {
    DiagnosticSeverity as RuleSeverity,
    SeverityOverrides,
} from './analyzer';

function psSingleQuoted(value: string): string {
    return `'${value.replace(/'/g, "''")}'`;
}

export function registerCommands(
    _context: vscode.ExtensionContext,
    bridge: PythonBridge,
    explorer: XlsmExplorer,
    _fsProvider: XlideFileSystemProvider,
    out: vscode.OutputChannel,
    vbaIndex: VbaSymbolIndex,
): vscode.Disposable[] {
    function log(msg: string): void {
        out.appendLine(msg);
    }

    // Builds a clickable Output-channel link to a module location. The link uses
    // the xlide-vba scheme with a triple-slash (empty authority) so the Output
    // panel's link detector recognises it, and an "#L<line>,<col>" fragment so
    // clicking reveals the exact problem location in the opened module.
    function moduleLocationLink(
        filePath: string,
        moduleName: string,
        line: number,
        column: number,
    ): string {
        const loc = encodeModuleUri(filePath, moduleName).with({
            fragment: `L${line},${column}`,
        });
        return loc.toString().replace(/^xlide-vba:\/(?!\/)/, 'xlide-vba://');
    }

    // Prints a formatted, scannable lint report with clickable per-problem
    // links. Leaves generous blank space above the report so it is easy to find
    // when the user opens the Output view.
    function printLintReport(
        filePath: string,
        moduleCount: number,
        problems: WorkbookLintProblem[],
        errorCount: number,
        warningCount: number,
        summary: WorkbookLintSummary,
    ): void {
        const name = path.basename(filePath);
        out.appendLine('');
        out.appendLine('');
        out.appendLine('');
        out.appendLine('========================================================================');
        out.appendLine(`  XLIDE Lint - ${name}`);
        out.appendLine(`  ${new Date().toLocaleString()}`);
        out.appendLine('========================================================================');
        out.appendLine('');

        if (problems.length === 0) {
            const suppressed = summary.suppressedCount > 0
                ? ` ${summary.suppressedCount} diagnostic(s) suppressed by XLIDE lint directives.`
                : '';
            out.appendLine(`  No unsuppressed problems found across ${moduleCount} module(s). Lint passed.${suppressed}`);
            out.appendLine('');
            return;
        }

        out.appendLine(
            `  ${problems.length} problem(s) in ${moduleCount} module(s): ` +
            `${errorCount} error(s), ${warningCount} warning(s).`,
        );
        out.appendLine(
            `  VBE compile-equivalent: ${summary.vbeCompileEquivalentCount}; ` +
            `XLIDE non-compile guidance/risk: ${summary.nonVbeCompileEquivalentCount}.`,
        );
        if (summary.suppressedCount > 0) {
            out.appendLine(`  Suppressed by XLIDE lint directives: ${summary.suppressedCount}.`);
        }
        out.appendLine(`  Categories: ${formatSummaryCounts(summary.byCategory)}.`);
        out.appendLine(`  Evidence: ${formatSummaryCounts(summary.byDiagnosticKind)}.`);
        out.appendLine('  (Click a location link to jump to the problem.)');
        out.appendLine('');

        let currentModule = '';
        for (const p of problems) {
            if (p.moduleName !== currentModule) {
                currentModule = p.moduleName;
                out.appendLine(`  ${currentModule} (${p.moduleType})`);
            }
            const sev = p.severity.toUpperCase().padEnd(11);
            const code = p.code ? ` [${p.code}]` : '';
            const link = moduleLocationLink(filePath, p.moduleName, p.line, p.column);
            out.appendLine(`    ${sev} ${p.message}${code}`);
            out.appendLine(`                ${link}`);
        }
        out.appendLine('');
    }

    function formatSummaryCounts<K extends string>(
        counts: Partial<Record<K, number>>,
    ): string {
        const entries = Object.entries(counts)
            .filter((entry): entry is [string, number] => typeof entry[1] === 'number' && entry[1] > 0)
            .sort(([a], [b]) => a.localeCompare(b));
        if (entries.length === 0) {
            return 'none';
        }
        return entries.map(([name, count]) => `${name} ${count}`).join(', ');
    }

    function shouldAttachToRunningExcel(): boolean {
        return vscode.workspace
            .getConfiguration('xlide')
            .get<boolean>('attachToRunningExcel', true);
    }

    function showRunMacroFailure(err: unknown): void {
        const raw = err instanceof Error ? err.message : String(err);
        const pipe = raw.indexOf('|');
        const code = pipe >= 0 ? raw.slice(0, pipe) : '';
        const message = pipe >= 0 ? raw.slice(pipe + 1) : raw;
        if (code === 'REOPEN_BLOCKED' || code === 'REOPEN_FAILED') {
            void vscode.window.showWarningMessage(`XLIDE: ${message}`);
            return;
        }
        void vscode.window.showErrorMessage(`XLIDE: Failed to run macro: ${message}`);
    }

    // Helper functions for Windows COM-based Excel operations
    function runWindowsExcel(filePath: string, attachToRunning: boolean, readOnly: boolean): void {
        const roFlag = readOnly ? '$true' : '$false';
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
            `  $workbook = $excel.Workbooks.Open($targetPath, 0, ${roFlag})`,
            '}',
            '$workbook.Activate()',
            "try { Add-Type -MemberDefinition '[DllImport(\"user32.dll\")] public static extern bool SetForegroundWindow(IntPtr hWnd); [DllImport(\"user32.dll\")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);' -Name XlideWin32 -Namespace XlideHelper } catch { }",
            '[XlideHelper.XlideWin32]::ShowWindow([IntPtr]$excel.Hwnd, 9)',
            '[XlideHelper.XlideWin32]::SetForegroundWindow([IntPtr]$excel.Hwnd)',
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

    function runWindowsExcelMacroReadOnly(filePath: string, macroName: string, attachToRunning: boolean): Promise<void> {
        const script = [
            '$ErrorActionPreference = "Stop"',
            'try {',
            `  $targetPath = ${psSingleQuoted(filePath)}`,
            `  $targetName = ${psSingleQuoted(path.basename(filePath))}`,
            `  $macroName = ${psSingleQuoted(macroName)}`,
            '  $excel = $null',
            '  $workbook = $null',
            `  $attachToRunning = ${attachToRunning ? '$true' : '$false'}`,
            '  if ($attachToRunning) {',
            '    try { $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application") } catch { }',
            '  }',
            '  if (-not $excel) {',
            '    $excel = New-Object -ComObject Excel.Application',
            '  }',
            '  $excel.Visible = $true',
            '  foreach ($wb in @($excel.Workbooks)) {',
            '    if (($wb.FullName -ieq $targetPath) -or ($wb.Name -ieq $targetName)) { $workbook = $wb; break }',
            '  }',
            '  if ($workbook) {',
            '    if (-not $workbook.ReadOnly) {',
            '      throw "REOPEN_BLOCKED|Workbook is already open for editing in Excel. Close it in Excel, then press F5 again so XLIDE can reopen the saved workbook before running the macro."',
            '    }',
            '    try {',
            '      $workbook.Close($false)',
            '      $workbook = $null',
            '    } catch {',
            '      throw ("REOPEN_FAILED|XLIDE could not close the existing read-only workbook before running the macro: " + $_.Exception.Message)',
            '    }',
            '  }',
            '  try {',
            '    $workbook = $excel.Workbooks.Open($targetPath, 0, $true)',
            '  } catch {',
            '    throw ("REOPEN_FAILED|XLIDE could not reopen the workbook. If it is open outside XLIDE, close it in Excel and try again: " + $_.Exception.Message)',
            '  }',
            '  $workbook.Activate()',
            '  try { Add-Type -MemberDefinition \'[DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd); [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);\' -Name XlideWin32 -Namespace XlideHelper } catch { }',
            '  [XlideHelper.XlideWin32]::ShowWindow([IntPtr]$excel.Hwnd, 9)',
            '  [XlideHelper.XlideWin32]::SetForegroundWindow([IntPtr]$excel.Hwnd)',
            '  $macroRef = "\'" + $workbook.Name + "\'!" + $macroName',
            '  try {',
            '    $excel.Run($macroRef)',
            '  } catch {',
            '    throw ("RUN_FAILED|XLIDE could not run the macro: " + $_.Exception.Message)',
            '  }',
            '  [Console]::Out.WriteLine("XLIDE_MACRO_OK")',
            '} catch {',
            '  [Console]::Error.WriteLine("XLIDE_MACRO_ERROR|" + $_.Exception.Message)',
            '  exit 1',
            '}',
        ].join('; ');

        log(`[runMacro] Running: ${macroName}`);
        log(`[runMacro] Script: ${script}`);
        return new Promise<void>((resolve, reject) => {
            const child = cp.spawn('powershell.exe', [
                '-NoProfile',
                '-ExecutionPolicy',
                'Bypass',
                '-Command',
                script,
            ]);
            const stderrLines: string[] = [];
            child.on('spawn', () => {
                log(`[runMacro] Spawned powershell.exe (pid=${child.pid ?? 'unknown'})`);
            });
            child.on('error', (err) => {
                log(`[runMacro] Error: ${err.message}`);
                reject(new Error(`RUN_FAILED|${err.message}`));
            });
            child.stdout?.on('data', (d: Buffer) => {
                const text = d.toString().trim();
                if (text) {
                    log(`[runMacro stdout] ${text}`);
                }
            });
            child.stderr?.on('data', (d: Buffer) => {
                for (const line of d.toString().split('\n')) {
                    const trimmed = line.trimEnd();
                    if (trimmed) {
                        stderrLines.push(trimmed);
                        log(`[runMacro stderr] ${trimmed}`);
                    }
                }
            });
            child.on('exit', (code, signal) => {
                log(`[runMacro] powershell exited with code=${code} signal=${signal ?? 'none'}`);
                if (code === 0) {
                    resolve();
                    return;
                }
                const sentinel = stderrLines.find((line) => line.includes('XLIDE_MACRO_ERROR|'));
                const message = sentinel
                    ? sentinel.slice(sentinel.indexOf('XLIDE_MACRO_ERROR|') + 'XLIDE_MACRO_ERROR|'.length)
                    : stderrLines.join('\n') || `PowerShell exited with code ${code}`;
                reject(new Error(message));
            });
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

    function diagnosticSeverityOverridesFromConfig(): SeverityOverrides {
        const optionExplicit = vscode.workspace
            .getConfiguration('xlide')
            .get<string>('diagnostics.optionExplicit', 'warning');
        return {
            optionExplicitMissing:
                optionExplicit === 'off' ? 'off' : (optionExplicit as RuleSeverity),
        };
    }

    async function lintActiveModule(): Promise<void> {
        const editor = vscode.window.activeTextEditor;
        if (!editor || editor.document.uri.scheme !== XLIDE_SCHEME) {
            vscode.window.showWarningMessage('XLIDE: Open a workbook VBA module to lint the current module.');
            return;
        }

        const { xlsmPath, moduleName } = decodeModuleUri(editor.document.uri);
        const source = editor.document.getText();
        const modules = await vbaIndex.getAllModules(xlsmPath);
        const current = modules.find(
            (mod) => mod.moduleName.toLowerCase() === moduleName.toLowerCase(),
        );
        const moduleKind = moduleKindFromType(current?.type);
        const moduleType = current?.type ?? 'standard';
        const project = buildVbaProjectIndex(modules, {
            moduleName,
            moduleKind,
            source,
        });
        const projectOptions = projectAnalysisOptionsForModule(
            project,
            moduleName,
            projectProcedureSignatures(project),
        );

        const result = lintVbaModuleSource({
            source,
            moduleName,
            moduleKind,
            documentType: current?.documentType,
            severities: diagnosticSeverityOverridesFromConfig(),
            ...projectOptions,
        });
        const problems = workbookProblemsForModule(
            moduleName,
            moduleType,
            source,
            result.diagnostics,
        ).sort((a, b) => {
            if (a.line !== b.line) { return a.line - b.line; }
            return a.column - b.column;
        });
        const errorCount = problems.filter((p) => p.severity === 'error').length;
        const warningCount = problems.filter((p) => p.severity === 'warning').length;
        const summary = summarizeWorkbookLintProblems(problems, result.suppressedCount);

        printLintReport(xlsmPath, 1, problems, errorCount, warningCount, summary);
        if (problems.length === 0) {
            vscode.window.showInformationMessage(
                `XLIDE: "${moduleName}" passed lint (no unsuppressed problems).`,
            );
        } else {
            vscode.window.showWarningMessage(
                `XLIDE: "${moduleName}" has ${errorCount} error(s) and ${warningCount} warning(s). See XLIDE Output.`,
            );
        }
    }

    async function showClassModuleReferences(node: XlideNode): Promise<void> {
        if (!node.moduleName || !node.filePath || node.isRemote) { return; }
        const originUri = encodeModuleUri(node.filePath, node.moduleName);
        const originDoc = await vscode.workspace.openTextDocument(originUri);
        await vscode.languages.setTextDocumentLanguage(originDoc, 'vba');
        const editor = await vscode.window.showTextDocument(originDoc, { preview: false });
        const origin = new vscode.Position(0, 0);
        editor.selection = new vscode.Selection(origin, origin);
        await vscode.commands.executeCommand('references-view.findReferences', originUri, origin);
    }

    function sameWorkbook(a: string, b: string): boolean {
        return process.platform === 'win32'
            ? a.toLowerCase() === b.toLowerCase()
            : a === b;
    }

    function applyOpenDocumentSources(
        modules: Awaited<ReturnType<VbaSymbolIndex['getAllModules']>>,
        xlsmPath: string,
    ): Awaited<ReturnType<VbaSymbolIndex['getAllModules']>> {
        const out = modules.map((mod) => ({ ...mod }));
        const byName = new Map(out.map((mod) => [mod.moduleName.toLowerCase(), mod]));
        for (const document of vscode.workspace.textDocuments) {
            if (document.uri.scheme !== XLIDE_SCHEME) { continue; }
            try {
                const decoded = decodeModuleUri(document.uri);
                if (!sameWorkbook(decoded.xlsmPath, xlsmPath)) { continue; }
                const mod = byName.get(decoded.moduleName.toLowerCase());
                if (!mod) { continue; }
                mod.source = document.getText();
            } catch {
                // Ignore non-standard xlide-vba URIs.
            }
        }
        return out;
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

        // Find all references to the procedure or class represented by a tree node
        vscode.commands.registerCommand('xlide.findReferences', async (node: XlideNode) => {
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
            await vscode.languages.setTextDocumentLanguage(doc, 'vba');
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

        // Add a new class module
        vscode.commands.registerCommand('xlide.newClassModule', async (node: XlideNode) => {
            if (node?.kind !== 'xlsm') { return; }
            const name = await vscode.window.showInputBox({
                prompt: 'New class module name',
                placeHolder: 'MyClass',
                validateInput: (v) =>
                    VBA_IDENTIFIER_NAME_RE.test(v) ? undefined : 'Class module names must be valid VBA identifiers',
            });
            if (!name) { return; }

            const stub = `Option Explicit\r\n\r\nPrivate Sub Class_Initialize()\r\n\r\nEnd Sub\r\n\r\nPrivate Sub Class_Terminate()\r\n\r\nEnd Sub\r\n`;
            try {
                await bridge.call('writeModule', {
                    path: node.filePath,
                    module: name,
                    source: stub,
                    kind: 'class',
                });
                explorer.refresh();
                const uri = encodeModuleUri(node.filePath, name);
                const doc = await vscode.workspace.openTextDocument(uri);
                await vscode.window.showTextDocument(doc, { preview: false });
                await vscode.languages.setTextDocumentLanguage(doc, 'vba');
            } catch (err) {
                vscode.window.showErrorMessage(`XLIDE: Failed to create class module: ${err}`);
            }
        }),

        // Rename a module
        vscode.commands.registerCommand('xlide.renameModule', async (node: XlideNode) => {
            if (!node?.moduleName) { return; }
            const validateInput = (v: string): string | undefined => {
                if (node.moduleType === 'class') {
                    return VBA_IDENTIFIER_NAME_RE.test(v)
                        ? undefined
                        : 'Class module names must be valid VBA identifiers';
                }
                return /^\w+$/.test(v) ? undefined : 'Module names must be alphanumeric';
            };
            const newName = await vscode.window.showInputBox({
                prompt: `Rename "${node.moduleName}" to`,
                value: node.moduleName,
                validateInput,
            });
            if (!newName || newName === node.moduleName) { return; }

            let moduleRenamed = false;
            try {
                if (node.moduleType === 'class') {
                    const modules = applyOpenDocumentSources(
                        await vbaIndex.getAllModules(node.filePath),
                        node.filePath,
                    );
                    const project = buildVbaProjectIndex(modules);
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
                            await vscode.languages.setTextDocumentLanguage(doc, 'vba');
                        }
                        const applied = await vscode.workspace.applyEdit(references.edit);
                        if (!applied) {
                            throw new Error('VS Code did not apply the class reference edits.');
                        }
                    }
                } else {
                    const result = await bridge.call<{ ok: boolean; signatureDropped: boolean }>(
                        'renameModule',
                        {
                            path: node.filePath,
                            module: node.moduleName,
                            newName,
                        },
                    );
                    moduleRenamed = true;
                    notifySignatureDropped(node.filePath, result.signatureDropped);
                    vbaIndex.invalidate(node.filePath);
                }
            } catch (err) {
                const prefix = moduleRenamed
                    ? 'XLIDE: Module was renamed, but reference updates failed'
                    : 'XLIDE: Rename failed';
                vscode.window.showErrorMessage(`${prefix}: ${err}`);
            } finally {
                if (moduleRenamed) {
                    explorer.refresh();
                }
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
                const result = await bridge.call<{ ok: boolean; signatureDropped: boolean }>(
                    'deleteModule',
                    {
                        path: node.filePath,
                        module: node.moduleName,
                    },
                );
                notifySignatureDropped(node.filePath, result.signatureDropped);
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
                const exportMode = normalizeExportMode(existingConfig.exportMode);
                const configuredFolder = existingConfig.exportFolder;

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

        // Import selected module files from the configured (or user-chosen) export folder
        vscode.commands.registerCommand('xlide.importModulesFromFolder', async (node: XlideNode) => {
            const filePath = resolveWorkbookPath(node);
            if (!filePath) { return; }

            try {
                const existingConfig = await readWorkbookRepoConfig(filePath);
                const configuredFolder = existingConfig.exportFolder;

                let importFolder: string;
                if (configuredFolder) {
                    importFolder = configuredFolder;
                } else {
                    const selected = await vscode.window.showOpenDialog({
                        canSelectFiles: false,
                        canSelectFolders: true,
                        canSelectMany: false,
                        openLabel: 'Select folder to import from',
                        defaultUri: vscode.Uri.file(path.dirname(filePath)),
                    });
                    if (!selected || selected.length === 0) { return; }
                    importFolder = selected[0].fsPath;
                }

                let entries: string[];
                try {
                    entries = await fs.promises.readdir(importFolder);
                } catch {
                    vscode.window.showErrorMessage(`XLIDE: Cannot read folder: ${importFolder}`);
                    return;
                }

                const moduleFiles = entries
                    .filter(e => /\.(bas|cls|frm)$/i.test(e))
                    .sort();

                if (moduleFiles.length === 0) {
                    vscode.window.showInformationMessage(
                        `XLIDE: No .bas/.cls/.frm files found in ${importFolder}`,
                    );
                    return;
                }

                // Fetch live module list so we know which names exist and their types.
                let liveModuleMap = new Map<string, string>();
                try {
                    const liveModules = await bridge.call<Array<{ name: string; type: string }>>(
                        'listModules', { path: filePath },
                    );
                    liveModuleMap = new Map(liveModules.map(m => [m.name, m.type]));
                } catch {
                    // Non-fatal: without live data we can't detect missing document modules.
                    log('[importModules] Warning: could not fetch live module list');
                }

                // UserForms always carry TWO GUIDs in VB_Base (type-lib + instance).
                // Class and document modules each have exactly one.
                // Document CLSIDs are the known Excel Workbook/Sheet/Chart GUIDs.
                const GUID_RE = /\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}/g;
                const DOCUMENT_CLSIDS = new Set([
                    '{00020819-0000-0000-C000-000000000046}', // Workbook
                    '{00020820-0000-0000-C000-000000000046}', // Worksheet
                    '{00020821-0000-0000-C000-000000000046}', // Chart
                ]);

                async function fileModuleSubtype(file: string): Promise<'userform' | 'document' | 'class' | null> {
                    try {
                        const head = (await fs.promises.readFile(
                            path.join(importFolder, file), 'utf8',
                        )).slice(0, 2000);
                        const vbBaseMatch = head.match(/Attribute\s+VB_Base\s*=\s*"([^"]*)"/i);
                        if (vbBaseMatch) {
                            const guids = vbBaseMatch[1].match(new RegExp(GUID_RE.source, 'g')) ?? [];
                            if (guids.length >= 2) { return 'userform'; }
                            if (guids.some(g => DOCUMENT_CLSIDS.has(g.toUpperCase().replace(/\{([^}]+)\}/, '{$1}')))) {
                                return 'document';
                            }
                            return 'class';
                        }
                        if (/Attribute\s+VB_PredeclaredId\s*=\s*True/i.test(head)) {
                            return 'document';
                        }
                        return null;
                    } catch {
                        return null;
                    }
                }

                type ImportItem = vscode.QuickPickItem & { file: string; isDocumentMissing: boolean; moduleKind: string };

                const importable: ImportItem[] = [];
                const unavailable: ImportItem[] = [];

                for (const file of moduleFiles) {
                    const moduleName = path.basename(file, path.extname(file));
                    const liveType = liveModuleMap.get(moduleName);
                    const isFormFile = /\.frm$/i.test(file);
                    const isClsFile = /\.cls$/i.test(file);

                    // Determine the source-of-truth subtype: live workbook wins; otherwise inspect the file header.
                    let subtype: 'userform' | 'document' | 'class' | 'standard';
                    if (liveType === 'userform' || liveType === 'document' || liveType === 'class' || liveType === 'standard') {
                        subtype = liveType;
                    } else if (isFormFile) {
                        subtype = 'userform';
                    } else if (isClsFile) {
                        const fileSub = await fileModuleSubtype(file);
                        if (fileSub === 'userform') {
                            subtype = 'userform';
                        } else if (fileSub === 'document') {
                            subtype = 'document';
                        } else {
                            subtype = 'class';
                        }
                    } else {
                        subtype = 'standard';
                    }

                    // Userforms and document modules cannot be created from scratch \u2014
                    // they must already exist in the live workbook to be importable.
                    const requiresExisting = subtype === 'document' || subtype === 'userform';
                    const existsInWorkbook = liveModuleMap.has(moduleName);

                    if (requiresExisting && !existsInWorkbook) {
                        const kindLabel = subtype === 'userform' ? 'UserForm' : 'Document module';
                        unavailable.push({
                            label: moduleName,
                            description: file,
                            detail: `${kindLabel} \u2014 not in this workbook, cannot create`,
                            picked: false,
                            file,
                            isDocumentMissing: true,
                            moduleKind: subtype,
                        });
                    } else {
                        let newDetail: string;
                        switch (subtype) {
                            case 'class':
                                newDetail = 'Will be created as a new class module';
                                break;
                            case 'userform':
                                newDetail = 'Will update existing UserForm code';
                                break;
                            case 'document':
                                newDetail = 'Will update existing document module code';
                                break;
                            default:
                                newDetail = 'Will be created as a new standard module';
                        }
                        importable.push({
                            label: moduleName,
                            description: existsInWorkbook ? file : `${file}  (new)`,
                            detail: existsInWorkbook ? undefined : newDetail,
                            picked: true,
                            file,
                            isDocumentMissing: false,
                            moduleKind: subtype,
                        });
                    }
                }

                const quickPickItems: (vscode.QuickPickItem | ImportItem)[] = [...importable];
                if (unavailable.length > 0) {
                    quickPickItems.push(
                        { label: 'Cannot import', kind: vscode.QuickPickItemKind.Separator } as vscode.QuickPickItem,
                        ...unavailable,
                    );
                }

                const picks = await vscode.window.showQuickPick(
                    quickPickItems as ImportItem[],
                    {
                        canPickMany: true,
                        title: `Import modules into ${path.basename(filePath)}`,
                        placeHolder: 'Select modules to import (all selected by default)',
                    },
                );
                if (!picks || picks.length === 0) { return; }

                let importedCount = 0;
                const errors: string[] = [];
                for (const pick of picks) {
                    if ((pick as ImportItem).isDocumentMissing) { continue; }
                    try {
                        const source = await fs.promises.readFile(
                            path.join(importFolder, (pick as ImportItem).file),
                            'utf8',
                        );
                        log(`[importModules] Importing ${pick.label} from ${(pick as ImportItem).file}`);
                        await bridge.call('writeModule', {
                            path: filePath,
                            module: pick.label,
                            source,
                            kind: (pick as ImportItem).moduleKind,
                        });
                        importedCount++;
                    } catch (err) {
                        const message = err instanceof Error ? err.message : String(err);
                        errors.push(`${pick.label}: ${message}`);
                        log(`[importModules] Error importing ${pick.label}: ${message}`);
                    }
                }

                explorer.refresh();

                if (errors.length > 0) {
                    void vscode.window.showWarningMessage(
                        `XLIDE: Imported ${importedCount} module(s), ${errors.length} failed. See XLIDE Output for details.`,
                        'View Output',
                    ).then(choice => {
                        if (choice === 'View Output') {
                            void vscode.commands.executeCommand('xlide.showOutput');
                        }
                    });
                } else {
                    vscode.window.showInformationMessage(
                        `XLIDE: Imported ${importedCount} module(s) into ${path.basename(filePath)}`,
                    );
                }
            } catch (err) {
                const message = err instanceof Error ? err.message : String(err);
                log(`[importModules] Error: ${message}`);
                vscode.window.showErrorMessage(`XLIDE: Import failed: ${message}`);
            }
        }),

        // Change the configured export folder for this workbook
        vscode.commands.registerCommand('xlide.changeRepoFolder', async (node: XlideNode) => {
            const filePath = resolveWorkbookPath(node);
            if (!filePath) { return; }

            try {
                const existingConfig = await readWorkbookRepoConfig(filePath);
                const currentFolder = existingConfig.exportFolder;
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
                const currentMode = normalizeExportMode(existingConfig.exportMode);
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

        // DEV: smoke test — verifies listModules + readModule against a workspace workbook
        vscode.commands.registerCommand('xlide.dev.smoke', async () => {
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
                        const msg = err instanceof Error ? err.message : String(err);
                        log(`[smoke] FAILED: ${msg}`);
                        vscode.window.showErrorMessage(`XLIDE Smoke FAILED: ${msg}`);
                    }
                },
            );
        }),

        // Validate the workbook's VBA project structure
        vscode.commands.registerCommand('xlide.validateWorkbook', async (node: XlideNode) => {
            const filePath = resolveWorkbookPath(node);
            if (!filePath) {
                vscode.window.showWarningMessage('XLIDE: No workbook selected to validate.');
                return;
            }
            const name = path.basename(filePath);
            await vscode.window.withProgress(
                { location: vscode.ProgressLocation.Notification, title: `XLIDE: Validating "${name}"...`, cancellable: false },
                async () => {
                    try {
                        const res = await bridge.call<{ issues: string[] }>('validateWorkbook', { path: filePath });
                        const issues = res.issues ?? [];
                        if (issues.length === 0) {
                            log(`[validate] "${name}": no issues`);
                            void vscode.window.showInformationMessage(`XLIDE: "${name}" passed validation (no issues).`);
                            return;
                        }
                        log(`[validate] "${name}": ${issues.length} issue(s):`);
                        for (const issue of issues) {
                            log(`[validate]   - ${issue}`);
                        }
                        void vscode.window.showWarningMessage(
                            `XLIDE: "${name}" has ${issues.length} validation issue(s). See XLIDE Output for details.`,
                        );
                    } catch (err) {
                        const msg = err instanceof Error ? err.message : String(err);
                        log(`[validate] FAILED: ${msg}`);
                        vscode.window.showErrorMessage(`XLIDE: Validation failed: ${msg}`);
                    }
                },
            );
        }),

        // Lint the active VBA module using the same source text the editor shows.
        vscode.commands.registerCommand('xlide.lintCurrentModule', async () => {
            try {
                await lintActiveModule();
            } catch (err) {
                const message = err instanceof Error ? err.message : String(err);
                log(`[lintCurrentModule] Error: ${message}`);
                vscode.window.showErrorMessage(`XLIDE: Failed to lint current module: ${message}`);
            }
        }),

        // Lint every VBA module in the workbook and print a clickable report
        vscode.commands.registerCommand('xlide.lintWorkbook', async (node: XlideNode) => {
            const filePath = resolveWorkbookPath(node);
            if (!filePath) {
                vscode.window.showWarningMessage('XLIDE: No workbook selected to lint.');
                return;
            }
            const name = path.basename(filePath);
            await vscode.window.withProgress(
                { location: vscode.ProgressLocation.Notification, title: `XLIDE: Linting "${name}"...`, cancellable: false },
                async () => {
                    try {
                        const result = await lintWorkbook(bridge, filePath);
                        printLintReport(
                            filePath,
                            result.moduleCount,
                            result.problems,
                            result.errorCount,
                            result.warningCount,
                            result.summary,
                        );
                        if (result.problems.length === 0) {
                            void vscode.window.showInformationMessage(
                                `XLIDE: "${name}" passed lint (no problems across ${result.moduleCount} module(s)).`,
                            );
                        } else {
                            void vscode.window.showWarningMessage(
                                `XLIDE: "${name}" has ${result.errorCount} error(s) and ${result.warningCount} warning(s). See XLIDE Output.`,
                            );
                        }
                    } catch (err) {
                        const msg = err instanceof Error ? err.message : String(err);
                        log(`[lint] FAILED: ${msg}`);
                        vscode.window.showErrorMessage(`XLIDE: Lint failed: ${msg}`);
                    }
                },
            );
        }),

        // Create a new, empty macro-enabled workbook
        vscode.commands.registerCommand('xlide.newWorkbook', async () => {
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
                    try {
                        await bridge.call<{ ok: boolean; path: string }>('createWorkbook', { path: filePath });
                        log(`[newWorkbook] Created "${filePath}"`);
                        explorer.refresh();
                        void vscode.window.showInformationMessage(`XLIDE: Created "${name}".`);
                    } catch (err) {
                        const msg = err instanceof Error ? err.message : String(err);
                        log(`[newWorkbook] FAILED: ${msg}`);
                        vscode.window.showErrorMessage(`XLIDE: Failed to create workbook: ${msg}`);
                    }
                },
            );
        }),

        // Open the workbook in Excel (editable)
        vscode.commands.registerCommand('xlide.openWorkbook', async (node: XlideNode) => {
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
        vscode.commands.registerCommand('xlide.openWorkbookReadOnly', async (node: XlideNode) => {
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
        vscode.commands.registerCommand('xlide.runMacroAtCursor', async () => {
            const editor = vscode.window.activeTextEditor;
            if (!editor || !editor.document.uri.scheme.startsWith(XLIDE_SCHEME)) {
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
                    await runWindowsExcelMacroReadOnly(xlsmPath, `${moduleName}.${currentProc}`, attachToRunning);
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
