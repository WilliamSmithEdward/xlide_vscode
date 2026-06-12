// Shared document-identity helpers for the VBA editor surfaces.
//
// Answers "what module is this TextDocument?" (name, module kind, standalone
// vs workbook-backed) and builds the live ProjectIndex for a document either
// from the shared workbook context or from the lone editor buffer.

import * as vscode from 'vscode';
import {
    XLIDE_SCHEME,
    decodeModuleUri,
} from './xlideFileSystem';
import { ProjectIndex, type ModuleSymbolKind } from './analyzer';
import { buildLiveVbaProjectIndexAsync } from './vbaProjectAnalysis';
import { VbaProjectIndexService } from './vbaProjectIndexService';

export function isStandaloneVbaDocument(document: vscode.TextDocument): boolean {
    return document.uri.scheme !== XLIDE_SCHEME && document.languageId === 'vba';
}

export function moduleNameFromDocument(document: vscode.TextDocument): string {
    if (document.uri.scheme === XLIDE_SCHEME) {
        try {
            return decodeModuleUri(document.uri).moduleName;
        } catch {
            /* fall through */
        }
    }
    const base = document.uri.path.split('/').pop() ?? 'Module';
    return base.replace(/\.[^.]+$/, '') || 'Module';
}

export function moduleKindFromDocument(document: vscode.TextDocument): ModuleSymbolKind {
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
    if (document.uri.scheme !== XLIDE_SCHEME) {
        return buildLiveVbaProjectIndexAsync([], {
            moduleName,
            moduleKind: moduleKindFromDocument(document),
            source,
        }, {
            cancelIfRequested: () => {
                if (token?.isCancellationRequested) {
                    throw new vscode.CancellationError();
                }
            },
        });
    }

    // The shared workbook context already folds in the open editors' text
    // (including this document) one changed module at a time.
    const decoded = decodeModuleUri(document.uri);
    const context = await projectIndexService.contextForWorkbook(decoded.xlsmPath, 'live');
    if (token?.isCancellationRequested) {
        throw new vscode.CancellationError();
    }
    return context.project;
}
