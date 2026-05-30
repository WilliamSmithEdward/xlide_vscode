import * as vscode from 'vscode';
import { PythonBridge } from './pythonBridge';
import { XlsmExplorer } from './xlsmExplorer';
import { XlideFileSystemProvider, encodeModuleUri, notifySignatureDropped } from './xlideFileSystem';
import {
    type ExportMode,
    exportWorkbookModules,
    setWorkbookExportMode,
} from './moduleDump';

// --------------------------------------------------------------------------
// Input types matching the inputSchema in package.json
// --------------------------------------------------------------------------

interface ListModulesInput { filePath: string; }
interface ListSubsInput    { filePath: string; moduleName: string; }
interface ReadModuleInput  { filePath: string; moduleName: string; }
interface WriteModuleInput { filePath: string; moduleName: string; source: string; }
interface RenameModuleInput { filePath: string; moduleName: string; newName: string; }
interface DeleteModuleInput { filePath: string; moduleName: string; }
interface ListSheetsInput  { filePath: string; }
interface GetWorkbookInfoInput { filePath: string; }
interface ValidateWorkbookInput { filePath: string; }
interface CreateWorkbookInput { filePath: string; }
interface ReadCellsInput   { filePath: string; sheet: string; range: string; }
interface ReadFormulasInput { filePath: string; sheet: string; range: string; }
interface WriteCellsInput  { filePath: string; sheet: string; startCell: string; data: unknown[][]; }
interface RunOpenpyxlInput { filePath: string; code: string; save?: boolean; }
interface ExportModulesInput { filePath: string; exportFolder?: string; exportMode?: ExportMode; }
interface ConfigureExportModeInput { filePath: string; exportMode: ExportMode; }

function textResult(value: string): vscode.LanguageModelToolResult {
    return new vscode.LanguageModelToolResult([
        new vscode.LanguageModelTextPart(value),
    ]);
}

export function registerAgentTools(
    _context: vscode.ExtensionContext,
    bridge: PythonBridge,
    explorer: XlsmExplorer,
    fsProvider: XlideFileSystemProvider,
): vscode.Disposable[] {
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
            async invoke(options, _token) {
                const modules = await bridge.call<Array<{ name: string; type: string }>>(
                    'listModules',
                    { path: options.input.filePath },
                );
                return textResult(JSON.stringify(modules, null, 2));
            },
        }),

        // ----------------------------------------------------------------
        // xlide_listSubs
        // ----------------------------------------------------------------
        vscode.lm.registerTool<ListSubsInput>('xlide_listSubs', {
            async invoke(options, _token) {
                const subs = await bridge.call<Array<{ name: string; kind: string; line: number }>>(
                    'listSubs',
                    { path: options.input.filePath, module: options.input.moduleName },
                );
                return textResult(JSON.stringify(subs, null, 2));
            },
        }),

        // ----------------------------------------------------------------
        // xlide_readModule
        // ----------------------------------------------------------------
        vscode.lm.registerTool<ReadModuleInput>('xlide_readModule', {
            async invoke(options, _token) {
                const result = await bridge.call<{ source: string }>(
                    'readModule',
                    { path: options.input.filePath, module: options.input.moduleName },
                );
                return textResult(result.source);
            },
        }),

        // ----------------------------------------------------------------
        // xlide_writeModule  (requires user confirmation)
        // ----------------------------------------------------------------
        vscode.lm.registerTool<WriteModuleInput>('xlide_writeModule', {
            async invoke(options, _token) {
                const { filePath, moduleName, source } = options.input;
                const result = await bridge.call<{ ok: boolean; signatureDropped: boolean }>('writeModule', {
                    path: filePath,
                    module: moduleName,
                    source,
                });
                notifySignatureDropped(filePath, result.signatureDropped);
                explorer.refresh();
                // Notify VS Code that the file changed so open editors reload
                const uri = encodeModuleUri(filePath, moduleName);
                fsProvider.notifyFileChanged(uri);
                return textResult(
                    `Module "${moduleName}" written successfully to "${filePath}".`,
                );
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
                const result = await bridge.call<{ ok: boolean; signatureDropped: boolean }>(
                    'renameModule', { path: filePath, module: moduleName, newName },
                );
                notifySignatureDropped(filePath, result.signatureDropped);
                explorer.refresh();
                return textResult(`Module "${moduleName}" renamed to "${newName}" in "${filePath}".`);
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
                const result = await bridge.call<{ ok: boolean; signatureDropped: boolean }>(
                    'deleteModule', { path: filePath, module: moduleName },
                );
                notifySignatureDropped(filePath, result.signatureDropped);
                explorer.refresh();
                return textResult(`Module "${moduleName}" deleted from "${filePath}".`);
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
            async invoke(options, _token) {
                const result = await bridge.call<{ sheets: Array<{ name: string; dimensions: string }> }>(
                    'listSheets',
                    { path: options.input.filePath },
                );
                return textResult(JSON.stringify(result.sheets, null, 2));
            },
        }),

        // ----------------------------------------------------------------
        // xlide_getWorkbookInfo
        // ----------------------------------------------------------------
        vscode.lm.registerTool<GetWorkbookInfoInput>('xlide_getWorkbookInfo', {
            async invoke(options, _token) {
                const result = await bridge.call<{
                    modules: Array<{ name: string; type: string }>;
                    sheets: Array<{ name: string; dimensions: string }>;
                    namedRanges: Array<{ name: string; ref: string }>;
                }>('getWorkbookInfo', { path: options.input.filePath });
                return textResult(JSON.stringify(result, null, 2));
            },
        }),

        // ----------------------------------------------------------------
        // xlide_validateWorkbook
        // ----------------------------------------------------------------
        vscode.lm.registerTool<ValidateWorkbookInput>('xlide_validateWorkbook', {
            async invoke(options, _token) {
                const result = await bridge.call<{ issues: string[] }>(
                    'validateWorkbook',
                    { path: options.input.filePath },
                );
                return textResult(JSON.stringify(result, null, 2));
            },
        }),

        // ----------------------------------------------------------------
        // xlide_createWorkbook
        // ----------------------------------------------------------------
        vscode.lm.registerTool<CreateWorkbookInput>('xlide_createWorkbook', {
            async invoke(options, _token) {
                const result = await bridge.call<{ ok: boolean; path: string }>(
                    'createWorkbook',
                    { path: options.input.filePath },
                );
                explorer.refresh();
                return textResult(JSON.stringify(result, null, 2));
            },
        }),

        // ----------------------------------------------------------------
        // xlide_readCells
        // ----------------------------------------------------------------
        vscode.lm.registerTool<ReadCellsInput>('xlide_readCells', {
            async invoke(options, _token) {
                const { filePath, sheet, range } = options.input;
                const result = await bridge.call<{ data: unknown[][] }>(
                    'readCells',
                    { path: filePath, sheet, range },
                );
                return textResult(JSON.stringify(result.data, null, 2));
            },
        }),

        // ----------------------------------------------------------------
        // xlide_readFormulas
        // ----------------------------------------------------------------
        vscode.lm.registerTool<ReadFormulasInput>('xlide_readFormulas', {
            async invoke(options, _token) {
                const { filePath, sheet, range } = options.input;
                const result = await bridge.call<{ data: unknown[][] }>(
                    'readFormulas',
                    { path: filePath, sheet, range },
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
                await bridge.call('writeCells', {
                    path: filePath,
                    sheet,
                    startCell,
                    data,
                });
                return textResult(
                    `Cells written to sheet "${sheet}" starting at "${startCell}" in "${filePath}".`,
                );
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
        // xlide_runOpenpyxl  (requires user confirmation)
        // ----------------------------------------------------------------
        vscode.lm.registerTool<RunOpenpyxlInput>('xlide_runOpenpyxl', {
            async invoke(options, _token) {
                const { filePath, code, save } = options.input;
                const result = await bridge.call<{ result: unknown; stdout: string }>(
                    'runOpenpyxl',
                    { path: filePath, code, save: save !== false },
                );
                const parts: string[] = [];
                if (result.stdout) { parts.push(`stdout:\n${result.stdout}`); }
                parts.push(`result: ${JSON.stringify(result.result, null, 2)}`);
                return textResult(parts.join('\n'));
            },
            async prepareInvocation(options, _token) {
                const { filePath, save } = options.input;
                const saveLabel = save === false ? 'without saving' : 'and save';
                return {
                    invocationMessage: `Running openpyxl code against "${filePath}"`,
                    confirmationMessages: {
                        title: 'Run openpyxl Code',
                        message: new vscode.MarkdownString(
                            `Execute Python/openpyxl code against \`${filePath}\` ${saveLabel}?\n\n` +
                            `The code runs with full openpyxl access to the workbook.`,
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
                const result = await exportWorkbookModules(bridge, { filePath, exportFolder, exportMode });
                return textResult(JSON.stringify(result, null, 2));
            },
            async prepareInvocation(options, _token) {
                const { filePath, exportFolder, exportMode } = options.input;
                return {
                    invocationMessage: `Exporting VBA modules for "${filePath}"`,
                    confirmationMessages: {
                        title: 'Export VBA Modules',
                        message: new vscode.MarkdownString(
                            `Export all modules for \`${filePath}\` using mode **${exportMode ?? 'trueUp'}**` +
                            `${exportFolder ? ` to folder \`${exportFolder}\`` : ' using configured folder'}` +
                            `?\n\nThis writes files and updates workbookname.extension.repo.json.`,
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
                const updated = await setWorkbookExportMode(filePath, exportMode);
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
                            `This updates workbookname.extension.repo.json beside the workbook.`,
                        ),
                    },
                };
            },
        }),
    ];
}
