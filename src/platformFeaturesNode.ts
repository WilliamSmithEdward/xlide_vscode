// The desktop platform: everything XLIDE can do when there is a shell, a
// local Office install, a git binary and a worker thread.
//
// The browser build never reaches this module - webBuild.js aliases it to
// platformFeaturesWeb.ts - which is what keeps child_process, worker_threads
// and the test host out of the web bundle.

import type * as vscode from 'vscode';
import { AnalysisWorkerClient } from './analysisWorkerClient';
import { registerAgentTools } from './agentTools';
import { registerOfficeAppCommands } from './commands/officeAppCommands';
import { registerModuleSyncCommands } from './commands/moduleSyncCommands';
import { registerSupportBundleCommands } from './commands/supportBundleCommands';
import { registerVbaTestCommands } from './commands/vbaTestCommands';
import type { CommandDeps } from './commands/shared';
import { GitChangeMarks, watchRepositoriesForMarks } from './gitChangeMarks';
import { gitModuleCompareDeps, registerGitCompareCommands } from './gitModuleCompare';
import type { ProjectEngine } from './projectEngine';
import type { ProjectExplorer } from './projectExplorer';
import { registerFormPreview } from './vbaFormPreview';
import { registerVb6FormDesigner } from './vb6FormDesigner';
import { cleanupStaleVbaTestHostTempDirsAsync } from './vbaTestTempFiles';
import type { VbaSymbolIndex } from './vbaSymbolIndex';
import type { XlideFileSystemProvider } from './xlideFileSystem';
import type { ChangeMarks, PlatformFeatures, TempDirCleanup } from './platformFeatures';

export const platformFeatures: PlatformFeatures = {
	name: 'desktop',

	createChangeMarks(bridge: ProjectEngine, log: (line: string) => void): ChangeMarks {
		return new GitChangeMarks(gitModuleCompareDeps(bridge), log);
	},

	watchRepositories(marks: ChangeMarks): vscode.Disposable[] {
		return watchRepositoriesForMarks(marks as GitChangeMarks);
	},

	createAnalysisWorker(scriptPath: string, log: (line: string) => void) {
		return new AnalysisWorkerClient(scriptPath, log);
	},

	registerDesigners(context: vscode.ExtensionContext, bridge: ProjectEngine): void {
		registerFormPreview(context, bridge);
		registerVb6FormDesigner(context, bridge);
	},

	registerAgentTools(
		context: vscode.ExtensionContext,
		bridge: ProjectEngine,
		explorer: ProjectExplorer,
		fsProvider: XlideFileSystemProvider,
		vbaIndex: VbaSymbolIndex,
	): vscode.Disposable[] {
		return registerAgentTools(context, bridge, explorer, fsProvider, vbaIndex);
	},

	registerPlatformCommands(deps: CommandDeps): vscode.Disposable[] {
		return [
			...registerOfficeAppCommands(deps),
			...registerModuleSyncCommands(deps),
			...registerSupportBundleCommands(deps),
			...registerVbaTestCommands(deps),
			...registerGitCompareCommands(deps.bridge),
		];
	},

	async cleanupStaleTempDirs(): Promise<TempDirCleanup | undefined> {
		return cleanupStaleVbaTestHostTempDirsAsync();
	},
};
