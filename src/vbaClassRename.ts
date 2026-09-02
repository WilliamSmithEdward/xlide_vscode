import * as vscode from 'vscode';
import { ProjectEngine } from './projectEngine';
import { notifySignatureDropped } from './xlideFileSystem';
import {
    projectClassReferenceLocations,
    type VbaNavigationModule,
} from './vbaNavigation';
import {
    ProjectIndex,
    type VbaProjectTypeName,
} from './analyzer';

export async function renameProjectClassModule(
    bridge: ProjectEngine,
    projectPath: string,
    oldName: string,
    newName: string,
): Promise<void> {
    const result = await bridge.call<{ ok: boolean; signatureDropped: boolean }>(
        'renameModule',
        { path: projectPath, module: oldName, newName },
    );
    notifySignatureDropped(projectPath, result.signatureDropped);
}

export function projectClassReferenceEdit(
    projectPath: string,
    byModule: Map<string, VbaNavigationModule>,
    project: ProjectIndex,
    oldName: string,
    definition: VbaProjectTypeName,
    newName: string,
): { edit: vscode.WorkspaceEdit; uris: vscode.Uri[]; count: number } {
    const edit = new vscode.WorkspaceEdit();
    const seenUris = new Map<string, vscode.Uri>();
    let count = 0;
    for (const loc of projectClassReferenceLocations(
        projectPath,
        byModule,
        project,
        oldName,
        definition,
        newName,
    )) {
        edit.replace(loc.uri, loc.range, newName);
        seenUris.set(loc.uri.toString(), loc.uri);
        count++;
    }
    return { edit, uris: [...seenUris.values()], count };
}
