// Workspace discovery for macro containers - the one place that decides
// which files "the workspace's Office files" means. macroContainerUi stays
// vscode-free for the pure-node surfaces; this module owns the vscode side.

import * as path from 'path';
import * as vscode from 'vscode';
import { MACRO_CONTAINER_GLOB } from './macroContainerUi';

/** Dependency folders no XLIDE surface should walk or list. */
export const MACRO_CONTAINER_FIND_EXCLUDES = '{**/node_modules/**,**/.venv/**,**/venv/**}';

/**
 * Lists the workspace's macro container files the way every XLIDE surface
 * agrees to see them: dependency folders excluded, Office owner-lock stubs
 * (`~$Name.xlsm`) dropped, file scheme only, sorted by path. The explorer,
 * the sidebar, and the agent's xlide_listWorkbooks all answer from here, so
 * an agent can never see a file the trees would hide.
 */
export async function findMacroContainerFiles(): Promise<vscode.Uri[]> {
    const uris = await vscode.workspace.findFiles(MACRO_CONTAINER_GLOB, MACRO_CONTAINER_FIND_EXCLUDES);
    return uris
        .filter((uri) => uri.scheme === 'file' && !path.basename(uri.fsPath).startsWith('~$'))
        .sort((left, right) => left.fsPath.localeCompare(right.fsPath));
}
