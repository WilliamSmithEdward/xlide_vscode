import * as vscode from 'vscode';
import { PythonBridge } from './pythonBridge';
import { XlsmExplorer } from './xlsmExplorer';
import { XlideFileSystemProvider, encodeModuleUri } from './xlideFileSystem';
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
interface ReadCellsInput   { filePath: string; sheet: string; range: string; }
interface WriteCellsInput  { filePath: string; sheet: string; startCell: string; data: unknown[][]; }
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
                await bridge.call('writeModule', {
                    path: filePath,
                    module: moduleName,
                    source,
                });
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
