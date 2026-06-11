import * as vscode from 'vscode';
import { PythonBridge } from './pythonBridge';
import { XlsmExplorer } from './xlsmExplorer';
import { XlideFileSystemProvider } from './xlideFileSystem';
import { VbaSymbolIndex } from './vbaSymbolIndex';
import { type CommandDeps } from './commands/shared';
import { registerAnalysisCommands } from './commands/analysisCommands';
import { registerMiscCommands } from './commands/miscCommands';
import { registerModuleSyncCommands } from './commands/moduleSyncCommands';
import { registerSupportBundleCommands } from './commands/supportBundleCommands';
import { registerVbaTestCommands } from './commands/vbaTestCommands';
import { registerWorkbookCrudCommands } from './commands/workbookCrudCommands';

/**
 * Composition root for the XLIDE command palette/explorer commands.
 * Each domain registers its own commands against the shared CommandDeps;
 * see src/commands/ for the per-domain modules.
 */
export function registerCommands(
    context: vscode.ExtensionContext,
    bridge: PythonBridge,
    explorer: XlsmExplorer,
    fsProvider: XlideFileSystemProvider,
    out: vscode.OutputChannel,
    vbaIndex: VbaSymbolIndex,
): vscode.Disposable[] {
    const deps: CommandDeps = { context, bridge, explorer, fsProvider, out, vbaIndex };

    return [
        ...registerMiscCommands(deps),
        ...registerWorkbookCrudCommands(deps),
        ...registerModuleSyncCommands(deps),
        ...registerSupportBundleCommands(deps),
        ...registerAnalysisCommands(deps),
        ...registerVbaTestCommands(deps),
    ];
}
