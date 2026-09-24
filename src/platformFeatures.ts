// What activation wires differently on a desktop and in a browser.
//
// XLIDE's core - the project tree, the module editor, analysis, refactoring -
// is the same code either way. A handful of features are not, because they
// need a shell, a local Office install, a git binary, or a worker thread:
//
//   - the git change marks, and the repository watching behind them
//   - the off-thread analysis worker
//   - the form designers, which read sidecar files from disk
//   - the language model tools
//   - mirroring the MCP server's edits, which listens on a local port
//   - commands that launch Office or shell out
//   - cleaning up temp directories a previous test-host run left behind
//
// Each is reached through this interface, so activate() reads the same on
// both platforms and the browser build simply gets implementations that do
// nothing. The browser build swaps the leaf module underneath (see
// webBuild.js), which is also what keeps child_process and worker_threads out
// of the web bundle: the desktop implementations are never imported there.

import type * as vscode from 'vscode';
import type { CommandDeps } from './commands/shared';
import type { GitChangeMarksSource } from './gitChangeMarks';
import type { ProjectEngine } from './projectEngine';
import type { ProjectExplorer } from './projectExplorer';
import type { AnalysisWorker } from './vbaProjectWideAnalysis';
import type { VbaSymbolIndex } from './vbaSymbolIndex';
import type { XlideFileSystemProvider } from './xlideFileSystem';
import { platformFeatures as leaf } from './platformFeaturesNode';

/** Modules that differ from the last commit, for the tree's M/A marks. */
export interface ChangeMarks extends GitChangeMarksSource {
	invalidate(projectPath: string): void;
	invalidateAll(): void;
	dispose(): void;
}

/** Left-over temp directories from a previous test-host run. */
export interface TempDirCleanup {
	scanned: number;
	deleted: number;
	failed: number;
}

export interface PlatformFeatures {
	/** A name for the output channel and the support bundle. */
	readonly name: 'desktop' | 'web';

	/**
	 * Marks for the tree. A virtual workspace has no git binary to ask, so
	 * the web build returns marks that are always absent rather than wrong.
	 */
	createChangeMarks(bridge: ProjectEngine, log: (line: string) => void): ChangeMarks;

	/** Follows commits made outside XLIDE, so the marks keep up. */
	watchRepositories(marks: ChangeMarks): vscode.Disposable[];

	/**
	 * Keeps full analysis passes off the extension-host thread. Undefined
	 * where there is no worker to start, which analysis already handles: it
	 * falls back to analyzing in the host.
	 */
	createAnalysisWorker(
		scriptPath: string,
		log: (line: string) => void,
	): AnalysisWorker | undefined;

	/** The UserForm preview and the VB6 form designer. */
	registerDesigners(context: vscode.ExtensionContext, bridge: ProjectEngine): void;

	/** The language model tools agents call. */
	registerAgentTools(
		context: vscode.ExtensionContext,
		bridge: ProjectEngine,
		explorer: ProjectExplorer,
		fsProvider: XlideFileSystemProvider,
		vbaIndex: VbaSymbolIndex,
	): vscode.Disposable[];

	/**
	 * Shows the edits the MCP server makes, which it reports to a loopback
	 * port, in the tree with a diff and Keep / Revert. A browser has no port
	 * to listen on.
	 */
	mirrorMcpEdits(
		context: vscode.ExtensionContext,
		bridge: ProjectEngine,
		explorer: ProjectExplorer,
		log: (line: string) => void,
	): vscode.Disposable[];

	/**
	 * Commands that need something only this platform has: launching Office,
	 * running the test host, exporting to a folder, comparing with git.
	 */
	registerPlatformCommands(deps: CommandDeps): vscode.Disposable[];

	/** Undefined where there is no temp directory to sweep. */
	cleanupStaleTempDirs(): Promise<TempDirCleanup | undefined>;
}

export const platformFeatures: PlatformFeatures = leaf;
