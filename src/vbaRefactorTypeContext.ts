// The project facts a refactoring needs to type a name, read from what the
// project index has already loaded (issue #1494). Shared by the lightbulb and
// the command palette, and free of the vscode API at run time so either can
// load it.

import type * as vscode from 'vscode';
import type { ExpressionTypeContext } from './analyzer';
import { hostTokensForProject } from './analyzer/host/hostLibraries';
import { hostObjectModelForTokens, hostTokenForFileName } from './analyzer/host/hostRegistry';
import { moduleIdentityKey } from './projectIdentity';
import { moduleLocationOfDocument } from './vbaDocumentLocation';
import type { VbaProjectIndexService } from './vbaProjectIndexService';

/**
 * What Extract Variable needs to type a project name: the project's types and
 * its host's model, from the project the index has already loaded. Nothing is
 * read on the caret path; before the project loads, the module stands alone.
 * `Sheet1` then extracts as `Dim sheet As Sheet1` with `Set`, where a module
 * alone gives Variant and a Let that raises 438 (issue #1494).
 */
export function extractVariableTypeContext(
    document: vscode.TextDocument,
    projects: VbaProjectIndexService | undefined,
): ExpressionTypeContext | undefined {
    const location = moduleLocationOfDocument(document);
    const project = location && projects?.cachedContextForProject(location.projectPath);
    if (!location || !project) {
        return undefined;
    }
    const metadata = project.moduleMetadata.get(moduleIdentityKey(location.moduleName));
    const model = hostObjectModelForTokens(hostTokensForProject(hostTokenForFileName(location.projectPath), project.references));
    const projectClassMembers = project.project.projectMemberSurfaces(location.moduleName, { includeClassValueFacts: false });
    return {
        moduleName: location.moduleName,
        moduleKind: metadata?.moduleKind,
        model,
        projectClassMembers,
        projectVisibleSymbols: project.project.visibleIdentifierSymbols(location.moduleName),
    };
}
