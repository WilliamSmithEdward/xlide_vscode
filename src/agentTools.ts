import * as vscode from 'vscode';
import { checkModuleContentToken, moduleContentToken } from './moduleContentToken';
import type { WorkbookAnalysisResult } from './vbaWorkbookAnalysis';
import * as fs from 'fs';
import * as path from 'path';
import { WorkbookEngine } from './workbookEngine';
import { XlsmExplorer } from './xlsmExplorer';
import { XlideFileSystemProvider } from './xlideFileSystem';
import { VbaSymbolIndex } from './vbaSymbolIndex';
import {
    deleteWorkbookModule,
    renameWorkbookModule,
    writeWorkbookModule,
    type WorkbookModuleOperationDeps,
} from './workbookModuleOperations';
import {
    exportWorkbookModules,
} from './moduleExport';
import { type ExportMode } from './workbookSettings';
import { setWorkbookModuleSyncExportMode } from './workbookModuleSyncSettings';
import { analyzeWorkbook } from './vbaWorkbookAnalysis';
import { executeVbaTestRun } from './vbaTestRunPipeline';
import { agentVbaTestArtifactPayloadFromPipeline } from './agentVbaTestArtifacts';
import {
    describeVbaTestSelection,
    summarizeVbaTestRun,
    type VbaTestSelectionOptions,
} from './vbaTestRunner';
import { formatChangeSummary, withWriteAudit } from './xlideWriteAudit';

// --------------------------------------------------------------------------
// Input types matching the inputSchema in package.json
// --------------------------------------------------------------------------

interface ListModulesInput { filePath: string; }
interface ListSubsInput    { filePath: string; moduleName: string; }
interface ReadModuleInput  { filePath: string; moduleName: string; startLine?: number; endLine?: number; }
interface WriteModuleInput { filePath: string; moduleName: string; source: string; expectedContentToken?: string; }
interface RenameModuleInput { filePath: string; moduleName: string; newName: string; }
interface DeleteModuleInput { filePath: string; moduleName: string; }
interface ListSheetsInput  { filePath: string; }
interface GetWorkbookInfoInput { filePath: string; }
interface ValidateWorkbookInput { filePath: string; }
interface AnalyzeWorkbookInput { filePath: string; moduleName?: string; }
interface RunVbaTestsInput {
    filePath: string;
    moduleName?: string;
    procedureName?: string;
    testIds?: string[];
    includeTags?: string[];
    excludeTags?: string[];
    failFast?: boolean;
    includeHostEvents?: boolean;
}
interface CreateWorkbookInput { filePath: string; }
interface ReadCellsInput   { filePath: string; sheet: string; range: string; }
interface ReadFormulasInput { filePath: string; sheet: string; range: string; }
interface WriteCellsInput  { filePath: string; sheet: string; startCell: string; data: unknown[][]; }
interface ExportModulesInput { filePath: string; exportFolder?: string; exportMode?: ExportMode; }
interface ConfigureExportModeInput { filePath: string; exportMode: ExportMode; }

function textResult(value: string): vscode.LanguageModelToolResult {
    return new vscode.LanguageModelToolResult([
        new vscode.LanguageModelTextPart(value),
    ]);
}

function vbaTestSelectionFromInput(input: RunVbaTestsInput): VbaTestSelectionOptions | undefined {
    const selection: VbaTestSelectionOptions = {
        moduleName: input.moduleName,
        procedureName: input.procedureName,
        testIds: input.testIds,
        includeTags: input.includeTags,
        excludeTags: input.excludeTags,
    };
    if (
        !selection.moduleName &&
        !selection.procedureName &&
        !selection.testIds?.length &&
        !selection.includeTags?.length &&
        !selection.excludeTags?.length
    ) {
        return undefined;
    }
    return selection;
}

export function registerAgentTools(
    _context: vscode.ExtensionContext,
    bridge: WorkbookEngine,
    explorer: XlsmExplorer,
    fsProvider: XlideFileSystemProvider,
    vbaIndex: VbaSymbolIndex,
): vscode.Disposable[] {
    const ops: WorkbookModuleOperationDeps = { bridge, explorer, fsProvider, vbaIndex };
    return [
        // ----------------------------------------------------------------
        // xlide_listWorkbooks
        // ----------------------------------------------------------------
        vscode.lm.registerTool<Record<string, never>>('xlide_listWorkbooks', {
            async invoke(_options, _token) {
                const uris = await vscode.workspace.findFiles('**/*.{xlsm,xlsb,xlam}');
                const files = uris.map((u) => u.fsPath).sort();
                return textResult(JSON.stringify(files, null, 2));
            },
        }),

        // ----------------------------------------------------------------
        // xlide_listModules
        // ----------------------------------------------------------------
        vscode.lm.registerTool<ListModulesInput>('xlide_listModules', {
            async invoke(options, token) {
                const modules = await bridge.call<Array<{ name: string; type: string }>>(
                    'listModules',
                    { path: options.input.filePath },
                    token,
                );
                return textResult(JSON.stringify(modules, null, 2));
            },
        }),

        // ----------------------------------------------------------------
        // xlide_listSubs
        // ----------------------------------------------------------------
        vscode.lm.registerTool<ListSubsInput>('xlide_listSubs', {
            async invoke(options, token) {
                const subs = await bridge.call<Array<{ name: string; kind: string; line: number }>>(
                    'listSubs',
                    { path: options.input.filePath, module: options.input.moduleName },
                    token,
                );
                return textResult(JSON.stringify(subs, null, 2));
            },
        }),

        // ----------------------------------------------------------------
        // xlide_readModule
        // ----------------------------------------------------------------
        vscode.lm.registerTool<ReadModuleInput>('xlide_readModule', {
            async invoke(options, token) {
                const { filePath, moduleName, startLine, endLine } = options.input;
                const result = await bridge.call<{ source: string }>(
                    'readModule',
                    { path: filePath, module: moduleName },
                    token,
                );
                const contentToken = moduleContentToken(result.source);
                const lines = result.source.split(/\r?\n/);
                // A window is over the WHOLE module, so the token still
                // describes what a later conditional write is checked against.
                const from = Math.max(1, startLine ?? 1);
                const to = Math.min(lines.length, endLine ?? lines.length);
                const windowed = startLine === undefined && endLine === undefined;
                const body = windowed ? result.source : lines.slice(from - 1, to).join('\n');
                const header = windowed
                    ? `contentToken: ${contentToken} (${lines.length} lines)`
                    : `contentToken: ${contentToken} (lines ${from}-${to} of ${lines.length})`;
                return textResult(`${header}\n${body}`);
            },
        }),

        // ----------------------------------------------------------------
        // xlide_writeModule  (requires user confirmation)
        // ----------------------------------------------------------------
        vscode.lm.registerTool<WriteModuleInput>('xlide_writeModule', {
            async invoke(options, _token) {
                const { filePath, moduleName, source, expectedContentToken } = options.input;
                if (expectedContentToken) {
                    const current = await bridge.call<{ source: string }>(
                        'readModule', { path: filePath, module: moduleName },
                    );
                    const stale = checkModuleContentToken(current.source, expectedContentToken, moduleName);
                    if (stale) {
                        return textResult(stale.message);
                    }
                }
                const { summary } = await withWriteAudit({
                    command: 'xlide_writeModule',
                    operation: 'write-module',
                    workbookPath: filePath,
                    moduleName,
                    failedSummary: 'Write module: 0 changed, 1 failed',
                }, async () => {
                    const result = await writeWorkbookModule(ops, { filePath, moduleName, source });
                    return {
                        result,
                        summary: formatChangeSummary({
                            operation: 'Write module',
                            changed: [moduleName],
                        }),
                    };
                });
                return textResult(`${summary}\nModule "${moduleName}" written successfully.`);
            },
            async prepareInvocation(options, _token) {
                const { filePath, moduleName } = options.input;
                return {
                    invocationMessage: `Writing VBA module "${moduleName}"`,
                    confirmationMessages: {
                        title: 'Write VBA Module',
                        message: new vscode.MarkdownString(
                            `Write changes to **${moduleName}** in \`${filePath}\`?\n\n` +
                            `This will overwrite the module source and save the workbook.`,
                        ),
                    },
                };
            },
        }),

        // ----------------------------------------------------------------
        // xlide_renameModule  (requires user confirmation)
        // ----------------------------------------------------------------
        vscode.lm.registerTool<RenameModuleInput>('xlide_renameModule', {
            async invoke(options, _token) {
                const { filePath, moduleName, newName } = options.input;
                const { summary } = await withWriteAudit({
                    command: 'xlide_renameModule',
                    operation: 'rename-module',
                    workbookPath: filePath,
                    moduleName,
                    failedSummary: 'Rename module: 0 changed, 1 failed',
                }, async () => {
                    const result = await renameWorkbookModule(ops, { filePath, moduleName, newName });
                    return {
                        result,
                        moduleName: newName,
                        summary: formatChangeSummary({
                            operation: 'Rename module',
                            changed: [`${moduleName} -> ${newName}`],
                        }),
                    };
                });
                return textResult(`${summary}\nModule "${moduleName}" renamed to "${newName}".`);
            },
            async prepareInvocation(options, _token) {
                const { filePath, moduleName, newName } = options.input;
                return {
                    invocationMessage: `Renaming module "${moduleName}" to "${newName}"`,
                    confirmationMessages: {
                        title: 'Rename VBA Module',
                        message: new vscode.MarkdownString(
                            `Rename module **${moduleName}** to **${newName}** in \`${filePath}\`?`,
                        ),
                    },
                };
            },
        }),

        // ----------------------------------------------------------------
        // xlide_deleteModule  (requires user confirmation)
        // ----------------------------------------------------------------
        vscode.lm.registerTool<DeleteModuleInput>('xlide_deleteModule', {
            async invoke(options, _token) {
                const { filePath, moduleName } = options.input;
                const { summary } = await withWriteAudit({
                    command: 'xlide_deleteModule',
                    operation: 'delete-module',
                    workbookPath: filePath,
                    moduleName,
                    failedSummary: 'Delete module: 0 changed, 1 failed',
                }, async () => {
                    const result = await deleteWorkbookModule(ops, { filePath, moduleName });
                    return {
                        result,
                        summary: formatChangeSummary({
                            operation: 'Delete module',
                            changed: [moduleName],
                        }),
                    };
                });
                return textResult(`${summary}\nModule "${moduleName}" deleted.`);
            },
            async prepareInvocation(options, _token) {
                const { filePath, moduleName } = options.input;
                return {
                    invocationMessage: `Deleting module "${moduleName}"`,
                    confirmationMessages: {
                        title: 'Delete VBA Module',
                        message: new vscode.MarkdownString(
                            `Permanently delete module **${moduleName}** from \`${filePath}\`?\n\n` +
                            `This cannot be undone.`,
                        ),
                    },
                };
            },
        }),

        // ----------------------------------------------------------------
        // xlide_listSheets
        // ----------------------------------------------------------------
        vscode.lm.registerTool<ListSheetsInput>('xlide_listSheets', {
            async invoke(options, token) {
                const result = await bridge.call<{ sheets: Array<{ name: string; dimensions: string }> }>(
                    'listSheets',
                    { path: options.input.filePath },
                    token,
                );
                return textResult(JSON.stringify(result.sheets, null, 2));
            },
        }),

        // ----------------------------------------------------------------
        // xlide_getWorkbookInfo
        // ----------------------------------------------------------------
        vscode.lm.registerTool<GetWorkbookInfoInput>('xlide_getWorkbookInfo', {
            async invoke(options, token) {
                const result = await bridge.call<{
                    modules: Array<{ name: string; type: string }>;
                    sheets: Array<{ name: string; dimensions: string }>;
                    namedRanges: Array<{ name: string; ref: string }>;
                }>('getWorkbookInfo', { path: options.input.filePath }, token);
                return textResult(JSON.stringify(result, null, 2));
            },
        }),

        // ----------------------------------------------------------------
        // xlide_validateWorkbook
        // ----------------------------------------------------------------
        vscode.lm.registerTool<ValidateWorkbookInput>('xlide_validateWorkbook', {
            async invoke(options, token) {
                const result = await bridge.call<{ issues: string[] }>(
                    'validateWorkbook',
                    { path: options.input.filePath },
                    token,
                );
                return textResult(JSON.stringify(result, null, 2));
            },
        }),

        // ----------------------------------------------------------------
        // xlide_analyzeWorkbook
        // ----------------------------------------------------------------
        vscode.lm.registerTool<AnalyzeWorkbookInput>('xlide_analyzeWorkbook', {
            async invoke(options, token) {
                const { filePath, moduleName } = options.input;
                const result = await analyzeWorkbook(bridge, filePath, { token });
                if (!moduleName) {
                    return textResult(JSON.stringify(result, null, 2));
                }
                // Checking one module you just edited should not mean reading
                // back every finding in the workbook.
                const wanted = moduleName.toLowerCase();
                const forModule = (p: { moduleName: string }) => p.moduleName.toLowerCase() === wanted;
                const problems = result.problems.filter(forModule);
                const scoped: WorkbookAnalysisResult = {
                    ...result,
                    moduleCount: 1,
                    problems,
                    suppressedProblems: result.suppressedProblems.filter(forModule),
                    errorCount: problems.filter((p) => p.severity === 'error').length,
                    warningCount: problems.filter((p) => p.severity === 'warning').length,
                };
                return textResult(JSON.stringify(scoped, null, 2));
            },
        }),

        // ----------------------------------------------------------------
        // xlide_runVbaTests
        // ----------------------------------------------------------------
        vscode.lm.registerTool<RunVbaTestsInput>('xlide_runVbaTests', {
            async invoke(options, _token) {
                const { filePath, failFast, includeHostEvents } = options.input;
                const selection = vbaTestSelectionFromInput(options.input);
                const result = await executeVbaTestRun(bridge, filePath, { selection, failFast });
                if (result.kind === 'blocked-support') {
                    return textResult(JSON.stringify({
                        ok: false,
                        blocked: true,
                        reason: 'test-support',
                        filePath,
                        support: result.support,
                    }, null, 2));
                }
                if (result.kind === 'blocked-com') {
                    return textResult(JSON.stringify({
                        ok: false,
                        blocked: true,
                        reason: 'excel-com',
                        filePath,
                        runtime: result.runtime,
                    }, null, 2));
                }

                const { execution } = result;
                const summary = summarizeVbaTestRun(execution.report);
                const ok = summary.failed === 0 &&
                    summary.timeout === 0 &&
                    summary.hostError === 0 &&
                    summary.xpass === 0;
                const artifacts = agentVbaTestArtifactPayloadFromPipeline(result.artifacts);
                return textResult(JSON.stringify({
                    ok,
                    summary,
                    artifacts,
                    report: execution.report,
                    ...(includeHostEvents ? { hostEvents: execution.hostEvents } : {}),
                }, null, 2));
            },
            async prepareInvocation(options, _token) {
                const { filePath, failFast } = options.input;
                const selection = vbaTestSelectionFromInput(options.input);
                const scope = describeVbaTestSelection(selection) || 'all tests';
                return {
                    invocationMessage: `Running XLIDE VBA tests for "${filePath}"`,
                    confirmationMessages: {
                        title: 'Run XLIDE VBA Tests',
                        message: new vscode.MarkdownString(
                            `Run **${scope}** in \`${filePath}\` through the XLIDE owned read-only Excel test host?` +
                            `${failFast ? '\n\nFail-fast is enabled.' : ''}`,
                        ),
                    },
                };
            },
        }),

        // ----------------------------------------------------------------
        // xlide_createWorkbook
        // ----------------------------------------------------------------
        vscode.lm.registerTool<CreateWorkbookInput>('xlide_createWorkbook', {
            async invoke(options, _token) {
                const { filePath } = options.input;
                const { result } = await withWriteAudit({
                    command: 'xlide_createWorkbook',
                    operation: 'create-workbook',
                    workbookPath: filePath,
                    failedSummary: 'Create workbook: 0 changed, 1 failed',
                }, async () => {
                    if (!path.isAbsolute(filePath)) {
                        // A relative path resolves against this process's cwd, which
                        // is not the user's workspace, so the existsSync overwrite guard
                        // below could check a different directory than the write.
                        throw new Error(
                            `xlide_createWorkbook requires an absolute filePath (got "${filePath}").`,
                        );
                    }
                    if (fs.existsSync(filePath)) {
                        throw new Error(
                            `Workbook already exists: "${filePath}". ` +
                            `xlide_createWorkbook does not overwrite existing workbooks - choose a different filePath.`,
                        );
                    }
                    const result = await bridge.call<{ ok: boolean; path: string }>(
                        'createWorkbook',
                        { path: filePath },
                    );
                    explorer.refresh();
                    return {
                        result,
                        summary: formatChangeSummary({
                            operation: 'Create workbook',
                            changed: [filePath],
                        }),
                    };
                });
                return textResult(JSON.stringify(result, null, 2));
            },
            async prepareInvocation(options, _token) {
                return {
                    invocationMessage: `Creating workbook "${options.input.filePath}"`,
                    confirmationMessages: {
                        title: 'Create New Workbook',
                        message: new vscode.MarkdownString(
                            `Create a new Excel workbook at \`${options.input.filePath}\`?`,
                        ),
                    },
                };
            },
        }),

        // ----------------------------------------------------------------
        // xlide_readCells
        // ----------------------------------------------------------------
        vscode.lm.registerTool<ReadCellsInput>('xlide_readCells', {
            async invoke(options, token) {
                const { filePath, sheet, range } = options.input;
                const result = await bridge.call<{ data: unknown[][] }>(
                    'readCells',
                    { path: filePath, sheet, range },
                    token,
                );
                return textResult(JSON.stringify(result.data, null, 2));
            },
        }),

        // ----------------------------------------------------------------
        // xlide_readFormulas
        // ----------------------------------------------------------------
        vscode.lm.registerTool<ReadFormulasInput>('xlide_readFormulas', {
            async invoke(options, token) {
                const { filePath, sheet, range } = options.input;
                const result = await bridge.call<{ data: unknown[][] }>(
                    'readFormulas',
                    { path: filePath, sheet, range },
                    token,
                );
                return textResult(JSON.stringify(result.data, null, 2));
            },
        }),

        // ----------------------------------------------------------------
        // xlide_writeCells  (requires user confirmation)
        // ----------------------------------------------------------------
        vscode.lm.registerTool<WriteCellsInput>('xlide_writeCells', {
            async invoke(options, _token) {
                const { filePath, sheet, startCell, data } = options.input;
                const { summary } = await withWriteAudit({
                    command: 'xlide_writeCells',
                    operation: 'write-cells',
                    workbookPath: filePath,
                    failedSummary: 'Write cells: 0 changed, 1 failed',
                }, async () => {
                    const result = await bridge.call('writeCells', {
                        path: filePath,
                        sheet,
                        startCell,
                        data,
                    });
                    return {
                        result,
                        summary: formatChangeSummary({
                            operation: 'Write cells',
                            changed: [`${sheet}!${startCell}`],
                        }),
                    };
                });
                return textResult(`${summary}\nCells written to sheet "${sheet}" starting at "${startCell}".`);
            },
            async prepareInvocation(options, _token) {
                const { filePath, sheet, startCell } = options.input;
                return {
                    invocationMessage: `Writing cells to "${sheet}" in "${filePath}"`,
                    confirmationMessages: {
                        title: 'Write Excel Cells',
                        message: new vscode.MarkdownString(
                            `Write data to sheet **${sheet}** starting at \`${startCell}\` in \`${filePath}\`?`,
                        ),
                    },
                };
            },
        }),
        // ----------------------------------------------------------------
        // xlide_exportModules  (requires user confirmation)
        // ----------------------------------------------------------------
        vscode.lm.registerTool<ExportModulesInput>('xlide_exportModules', {
            async invoke(options, _token) {
                const { filePath, exportFolder, exportMode } = options.input;
                const { result, summary } = await withWriteAudit({
                    command: 'xlide_exportModules',
                    operation: 'export-modules',
                    workbookPath: filePath,
                    targetPath: exportFolder,
                    failedSummary: 'Export modules: 0 changed, 1 failed',
                }, async () => {
                    const result = await exportWorkbookModules(bridge, { filePath, exportFolder, exportMode });
                    return {
                        result,
                        targetPath: result.exportFolder,
                        summary: formatChangeSummary({
                            operation: 'Export modules',
                            changed: result.writtenFiles,
                            removed: result.removedFiles,
                        }),
                    };
                });
                return textResult(JSON.stringify({ ...result, changeSummary: summary }, null, 2));
            },
            async prepareInvocation(options, _token) {
                const { filePath, exportFolder, exportMode } = options.input;
                return {
                    invocationMessage: `Exporting VBA modules for "${filePath}"`,
                    confirmationMessages: {
                        title: 'Export VBA Modules',
                        message: new vscode.MarkdownString(
                            `Export all modules for \`${filePath}\` using mode **${exportMode ?? 'exportAll'}**` +
                            `${exportFolder ? ` to folder \`${exportFolder}\`` : ' using configured folder'}` +
                            `?\n\nThis writes files and updates <workbook>.xlide_settings.json.`,
                        ),
                    },
                };
            },
        }),

        // ----------------------------------------------------------------
        // xlide_configureExportMode  (requires user confirmation)
        // ----------------------------------------------------------------
        vscode.lm.registerTool<ConfigureExportModeInput>('xlide_configureExportMode', {
            async invoke(options, _token) {
                const { filePath, exportMode } = options.input;
                const updated = await setWorkbookModuleSyncExportMode(filePath, exportMode);
                return textResult(JSON.stringify({ filePath, ...updated }, null, 2));
            },
            async prepareInvocation(options, _token) {
                const { filePath, exportMode } = options.input;
                return {
                    invocationMessage: `Configuring export mode for "${filePath}"`,
                    confirmationMessages: {
                        title: 'Configure Export Mode',
                        message: new vscode.MarkdownString(
                            `Set export mode for \`${filePath}\` to **${exportMode}**?\n\n` +
                            `This updates <workbook>.xlide_settings.json beside the workbook.`,
                        ),
                    },
                };
            },
        }),
    ];
}
