// Which VB6 project a file on disk belongs to.
//
// A workbook module announces its container in its URI; a VB6 module is a
// plain `.bas`/`.cls`/`.frm` file, and only the `.vbp` that names it says
// which project it is part of. This keeps that answer ready synchronously,
// the way the language surfaces need it, from the workspace's discovered
// projects and the engine's module listing for each. One file, one owner:
// the first manifest that names it wins, and a change to any `.vbp`
// rebuilds the whole map.

import * as vscode from 'vscode';
import type { WorkbookEngine } from './workbookEngine';
import { findMacroContainerFiles } from './macroContainerDiscovery';
import { isVb6ProjectPath } from './macroContainerUi';
import { workbookIdentityKey } from './workbookIdentity';

export interface Vb6ModuleOwner {
	/** The absolute path of the `.vbp`. */
	vbpPath: string;
	/** The module's name, from its `VB_Name` attribute. */
	moduleName: string;
	/** The module's kind as the engine lists it. */
	moduleType: string;
}

interface Vb6ProjectListing {
	vbpPath: string;
	modules: Array<{ name: string; type: string; filePath?: string }>;
}

/** The owner map one listing set implies: file identity key -> owner. */
export function ownersFromListings(listings: readonly Vb6ProjectListing[]): Map<string, Vb6ModuleOwner> {
	const owners = new Map<string, Vb6ModuleOwner>();
	for (const listing of listings) {
		for (const module of listing.modules) {
			if (!module.filePath) {
				continue;
			}
			const key = workbookIdentityKey(module.filePath);
			if (owners.has(key)) {
				continue;
			}
			owners.set(key, { vbpPath: listing.vbpPath, moduleName: module.name, moduleType: module.type });
		}
	}
	return owners;
}

let owners = new Map<string, Vb6ModuleOwner>();
let loading: Promise<void> | undefined;
let loader: (() => Promise<void>) | undefined;
const changeEmitter = new vscode.EventEmitter<void>();

/** Fires after the owner map is rebuilt. */
export const onDidChangeVb6Projects: vscode.Event<void> = changeEmitter.event;

/** The project a file belongs to, from the last completed load; undefined when none. */
export function vb6ModuleOwnerOf(fsPath: string): Vb6ModuleOwner | undefined {
	return owners.get(workbookIdentityKey(fsPath));
}

/** Resolves once the current load (if any) has landed. */
export function ensureVb6ProjectsLoaded(): Promise<void> {
	if (loading) {
		return loading;
	}
	if (owners.size === 0 && loader) {
		return loader();
	}
	return Promise.resolve();
}

/** Test hook: install an owner map directly. */
export function setVb6ModuleOwnersForTests(map: Map<string, Vb6ModuleOwner>): void {
	owners = map;
	changeEmitter.fire();
}

/**
 * Wires discovery: the workspace's `.vbp` files are listed through the
 * engine at activation and again whenever a manifest appears, changes, or
 * goes away. Listing failures leave the project out rather than failing
 * the map - a broken manifest is the tree's problem to report.
 */
export function registerVb6ProjectLocator(
	context: vscode.ExtensionContext,
	bridge: Pick<WorkbookEngine, 'call'>,
	log?: (line: string) => void,
): void {
	loader = async (): Promise<void> => {
		if (loading) {
			return loading;
		}
		const load = (async () => {
			const uris = (await findMacroContainerFiles()).filter((uri) => isVb6ProjectPath(uri.fsPath));
			const listings: Vb6ProjectListing[] = [];
			for (const uri of uris) {
				try {
					const modules = await bridge.call<Vb6ProjectListing['modules']>('listModules', { path: uri.fsPath });
					listings.push({ vbpPath: uri.fsPath, modules });
				} catch (err) {
					log?.(`[vb6] ${uri.fsPath}: ${err instanceof Error ? err.message : String(err)}`);
				}
			}
			owners = ownersFromListings(listings);
			changeEmitter.fire();
		})().finally(() => {
			if (loading === load) {
				loading = undefined;
			}
		});
		loading = load;
		return load;
	};
	const watcher = vscode.workspace.createFileSystemWatcher('**/*.vbp');
	const reload = (): void => { void loader?.(); };
	context.subscriptions.push(
		watcher,
		watcher.onDidCreate(reload),
		watcher.onDidChange(reload),
		watcher.onDidDelete(reload),
		vscode.workspace.onDidChangeWorkspaceFolders(reload),
		changeEmitter,
	);
	void loader();
}
