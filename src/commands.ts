import * as vscode from 'vscode';
import { ProjectEngine } from './projectEngine';
import { ProjectExplorer } from './projectExplorer';
import { XlideFileSystemProvider } from './xlideFileSystem';
import { VbaSymbolIndex } from './vbaSymbolIndex';
import { type CommandDeps } from './commands/shared';
import { registerAnalysisCommands } from './commands/analysisCommands';
import { registerMiscCommands } from './commands/miscCommands';
import { registerProjectCrudCommands } from './commands/projectCrudCommands';
import { registerRefactorCommands } from './commands/refactorCommands';
import { registerFormatCommands } from './commands/formatCommands';
import { platformFeatures } from './platformFeatures';

/**
 * Composition root for the XLIDE command palette/explorer commands.
 * Each domain registers its own commands against the shared CommandDeps;
 * see src/commands/ for the per-domain modules.
 *
 * The groups that need a shell, a local Office install, a git binary or a
 * real filesystem come from platformFeatures, so the browser build never
 * imports them (see src/platformFeatures.ts).
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
        ...registerAnalysisCommands(deps),
        ...registerRefactorCommands(deps),
        ...registerFormatCommands(deps),
        ...platformFeatures.registerPlatformCommands(deps),
    ];
}
