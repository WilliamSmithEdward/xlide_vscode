// Workspace loader for external VBA documentation metadata files.
//
// Discovers `*.vbref.xml` (configurable) metadata files anywhere in the open
// workspace, parses them with the pure analyzer parser, and maintains a live
// DocRegistry that the hover and signature-help providers consult. The registry
// is the "developer-defined overrides the curated library" layer: a developer can
// document any symbol - including host members and runtime functions - and have
// those descriptions appear in tooltips.
//
// Files are reloaded automatically when created, changed, or deleted. See
// user_guides/vba-doc-comments.md.

import * as vscode from 'vscode';
import { DocRegistry, parseMetadataFile } from './analyzer';
import {
	xlideDocsEnabledFromConfig,
	xlideDocsMetadataGlobFromConfig,
} from './globalSettings';
import { startPerformanceTrace } from './performanceTrace';

const EXCLUDE_GLOB = '**/node_modules/**';
const RELOAD_DEBOUNCE_MS = 250;

/** Reads the configured metadata glob, falling back to the default. */
function metadataGlob(): string {
	return xlideDocsMetadataGlobFromConfig(vscode.workspace.getConfiguration('xlide')).value;
}

/** True when the documentation feature is enabled in settings. */
function docsEnabled(): boolean {
	return xlideDocsEnabledFromConfig(vscode.workspace.getConfiguration('xlide')).value;
}

/**
 * Loads, caches, and watches external documentation metadata files for the
 * workspace, exposing a {@link DocRegistry} resolvers can read at any time.
 */
export class DocMetadataLoader {
	private readonly _registry = new DocRegistry();
	private _watcher: vscode.FileSystemWatcher | undefined;
	private _reloadTimer: ReturnType<typeof setTimeout> | undefined;

	/** The live registry. Empty until the first load completes. */
	get registry(): DocRegistry {
		return this._registry;
	}

	/** Performs the initial load and installs the file watcher. */
	async start(context: vscode.ExtensionContext): Promise<void> {
		await this.reload();
		this._installWatcher();
		context.subscriptions.push({
			dispose: (): void => this.dispose(),
		});
		context.subscriptions.push(
			vscode.workspace.onDidChangeConfiguration((e) => {
				if (
					e.affectsConfiguration('xlide.docs.metadataGlob') ||
					e.affectsConfiguration('xlide.docs.enabled')
				) {
					this._installWatcher();
					void this.reload();
				}
			}),
		);
	}

	/** Re-reads every metadata file and rebuilds the registry. */
	async reload(): Promise<void> {
		const trace = startPerformanceTrace('docs.reload');
		this._registry.clear();
		if (!docsEnabled()) {
			trace.end('ok', 'disabled');
			return;
		}
		let uris: vscode.Uri[];
		try {
			uris = await vscode.workspace.findFiles(metadataGlob(), EXCLUDE_GLOB);
		} catch {
			trace.end('failed', 'findFiles');
			return;
		}
		let loadedCount = 0;
		for (const uri of uris) {
			try {
				const bytes = await vscode.workspace.fs.readFile(uri);
				const text = Buffer.from(bytes).toString('utf8');
				this._registry.add(parseMetadataFile(text));
				loadedCount++;
			} catch {
				// Skip files we cannot read or parse; others still load.
			}
		}
		trace.end('ok', `${loadedCount}/${uris.length} files`);
	}

	private _installWatcher(): void {
		this._watcher?.dispose();
		this._watcher = undefined;
		if (!docsEnabled()) {
			return;
		}
		const watcher = vscode.workspace.createFileSystemWatcher(metadataGlob());
		const trigger = (): void => this._scheduleReload();
		watcher.onDidCreate(trigger);
		watcher.onDidChange(trigger);
		watcher.onDidDelete(trigger);
		this._watcher = watcher;
	}

	private _scheduleReload(): void {
		if (this._reloadTimer) {
			clearTimeout(this._reloadTimer);
		}
		this._reloadTimer = setTimeout(() => {
			this._reloadTimer = undefined;
			void this.reload();
		}, RELOAD_DEBOUNCE_MS);
	}

	/** Releases the watcher and any pending reload timer. */
	dispose(): void {
		if (this._reloadTimer) {
			clearTimeout(this._reloadTimer);
			this._reloadTimer = undefined;
		}
		this._watcher?.dispose();
		this._watcher = undefined;
	}
}
