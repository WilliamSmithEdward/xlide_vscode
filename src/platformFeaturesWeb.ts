// The web platform: what is left when there is no shell, no local Office, no
// git binary and no worker thread.
//
// Everything here is deliberately absent rather than approximated. A tree
// showing change marks computed without git, or an "Open in Excel" that
// silently did nothing, would be worse than not offering the feature: the
// commands are not contributed to the browser build's package.json either, so
// they do not appear in the palette at all.

import * as vscode from 'vscode';
import type { CommandDeps } from './commands/shared';
import type { GitProjectMarks } from './gitChangeMarks';
import type { ProjectEngine } from './projectEngine';
import type { ChangeMarks, PlatformFeatures, TempDirCleanup } from './platformFeatures';

/**
 * Marks that are always absent. `marksFor` returning undefined is the same
 * answer the desktop gives for a file git has no say over, so the tree
 * already knows how to draw it: with no mark.
 */
class NoChangeMarks implements ChangeMarks {
	private readonly _emitter = new vscode.EventEmitter<string>();
	readonly onDidChange = this._emitter.event;

	marksFor(_projectPath: string): GitProjectMarks | undefined {
		return undefined;
	}

	invalidate(_projectPath: string): void {
		/* nothing is cached */
	}

	invalidateAll(): void {
		/* nothing is cached */
	}

	dispose(): void {
		this._emitter.dispose();
	}
}

export const platformFeatures: PlatformFeatures = {
	name: 'web',

	createChangeMarks(_bridge: ProjectEngine, _log: (line: string) => void): ChangeMarks {
		return new NoChangeMarks();
	},

	watchRepositories(_marks: ChangeMarks): vscode.Disposable[] {
		return [];
	},

	createAnalysisWorker(_scriptPath: string, _log: (line: string) => void) {
		// No worker_threads in a browser. Analysis already falls back to
		// running in the extension host when there is no worker, which is
		// slower on a large project but produces the same diagnostics.
		return undefined;
	},

	registerDesigners(_context: vscode.ExtensionContext, _bridge: ProjectEngine): void {
		/* the designers read sidecar files from disk */
	},

	registerAgentTools(): vscode.Disposable[] {
		return [];
	},

	mirrorMcpEdits(): vscode.Disposable[] {
		// No port to listen on, and no MCP server beside a browser to report.
		return [];
	},

	registerPlatformCommands(_deps: CommandDeps): vscode.Disposable[] {
		return [];
	},

	async cleanupStaleTempDirs(): Promise<TempDirCleanup | undefined> {
		return undefined;
	},
};
