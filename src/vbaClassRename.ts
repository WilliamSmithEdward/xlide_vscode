import * as vscode from 'vscode';
import { PythonBridge } from './pythonBridge';
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
    bridge: PythonBridge,
    xlsmPath: string,
    oldName: string,
    newName: string,
): Promise<void> {
    const result = await bridge.call<{ ok: boolean; signatureDropped: boolean }>(
        'renameModule',
        { path: xlsmPath, module: oldName, newName },
    );
    notifySignatureDropped(xlsmPath, result.signatureDropped);
}

export function projectClassReferenceEdit(
    xlsmPath: string,
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
        xlsmPath,
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
