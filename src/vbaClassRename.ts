import * as vscode from 'vscode';
import {
    projectClassReferenceLocations,
    type VbaNavigationModule,
} from './vbaNavigation';
import {
    ProjectIndex,
    type VbaProjectTypeName,
} from './analyzer';

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
