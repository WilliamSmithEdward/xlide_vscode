// Shared document-identity helpers for the VBA editor surfaces.
//
// Answers "what module is this TextDocument?" (name, module kind, standalone
// vs project-backed) and builds the live ProjectIndex for a document either
// from the shared project context or from the lone editor buffer. A
// project module is backed by its container; a VB6 module file is backed
// by the `.vbp` that names it; a loose file nobody claims stands alone.

import * as vscode from 'vscode';
import { XLIDE_SCHEME } from './xlideFileSystem';
import { ProjectIndex, type ModuleSymbolKind } from './analyzer';
import { buildLiveVbaProjectIndexAsync, moduleKindFromType } from './vbaProjectAnalysis';
import { VbaProjectIndexService } from './vbaProjectIndexService';
import { analysisSourceForDocument, moduleLocationOfDocument } from './vbaDocumentLocation';

export { analysisSourceForDocument, moduleLocationOfDocument } from './vbaDocumentLocation';

/** A file on disk that no project claims: analyzed as a module on its own. */
export function isStandaloneVbaDocument(document: vscode.TextDocument): boolean {
    return document.uri.scheme !== XLIDE_SCHEME
        && document.languageId === 'vba'
        && moduleLocationOfDocument(document) === undefined;
}

export function moduleNameFromDocument(document: vscode.TextDocument): string {
    const location = moduleLocationOfDocument(document);
    if (location) {
        return location.moduleName;
    }
    const base = document.uri.path.split('/').pop() ?? 'Module';
    return base.replace(/\.[^.]+$/, '') || 'Module';
}

export function moduleKindFromDocument(document: vscode.TextDocument): ModuleSymbolKind {
    const location = moduleLocationOfDocument(document);
    if (location?.moduleType) {
        return moduleKindFromType(location.moduleType);
    }
    const fileName = document.uri.path.split('/').pop() ?? '';
    if (/\.cls$/i.test(fileName)) {
        return 'class';
    }
    if (/\.frm$/i.test(fileName)) {
        return 'userform';
    }
    return 'standard';
}

export async function liveProjectIndexForDocument(
    projectIndexService: VbaProjectIndexService,
    document: vscode.TextDocument,
    source: string,
    moduleName: string,
    token?: vscode.CancellationToken,
): Promise<ProjectIndex> {
    const location = moduleLocationOfDocument(document);
    if (!location) {
        return buildLiveVbaProjectIndexAsync([], {
            moduleName,
            moduleKind: moduleKindFromDocument(document),
            source: source === document.getText() ? analysisSourceForDocument(document) : source,
        }, {
            cancelIfRequested: () => {
                if (token?.isCancellationRequested) {
                    throw new vscode.CancellationError();
                }
            },
        });
    }

    // The shared project context already folds in the open editors' text
    // (including this document) one changed module at a time.
    const context = await projectIndexService.contextForProject(location.projectPath, 'live');
    if (token?.isCancellationRequested) {
        throw new vscode.CancellationError();
    }
    return context.project;
}
