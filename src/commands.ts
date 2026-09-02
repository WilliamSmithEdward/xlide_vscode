import * as vscode from 'vscode';
import { ProjectEngine } from './projectEngine';
import { ProjectExplorer } from './projectExplorer';
import { XlideFileSystemProvider } from './xlideFileSystem';
import { VbaSymbolIndex } from './vbaSymbolIndex';
import { type CommandDeps } from './commands/shared';
import { registerAnalysisCommands } from './commands/analysisCommands';
import { registerMiscCommands } from './commands/miscCommands';
import { registerModuleSyncCommands } from './commands/moduleSyncCommands';
import { registerSupportBundleCommands } from './commands/supportBundleCommands';
import { registerVbaTestCommands } from './commands/vbaTestCommands';
import { registerProjectCrudCommands } from './commands/projectCrudCommands';

/**
 * Composition root for the XLIDE command palette/explorer commands.
 * Each domain registers its own commands against the shared CommandDeps;
 * see src/commands/ for the per-domain modules.
 */
export function registerCommands(
    context: vscode.ExtensionContext,
    bridge: ProjectEngine,
    explorer: ProjectExplorer,
    fsProvider: XlideFileSystemProvider,
    out: vscode.OutputChannel,
    vbaIndex: VbaSymbolIndex,
): vscode.Disposable[] {
    const deps: CommandDeps = { context, bridge, explorer, fsProvider, out, vbaIndex };

    return [
        ...registerMiscCommands(deps),
        ...registerProjectCrudCommands(deps),
        ...registerModuleSyncCommands(deps),
        ...registerSupportBundleCommands(deps),
        ...registerAnalysisCommands(deps),
        ...registerVbaTestCommands(deps),
    ];
}
