// Composition root for the VBA language tooling: constructs the shared
// symbol index and project-index service, then wires up the subsystem
// modules (live diagnostics, typing automation, completion/hover/signature,
// navigation, code actions, and semantic tokens).

import * as vscode from 'vscode';
import { PythonBridge } from './pythonBridge';
import {
    XLIDE_SCHEME,
    XLIDE_VBA_LANGUAGE_ID,
    decodeModuleUri,
} from './xlideFileSystem';
import {
    VbaDefinitionProvider,
    VbaDocumentSymbolProvider,
    VbaReferenceProvider,
    VbaRenameProvider,
    VbaWorkspaceSymbolProvider,
} from './vbaNavigationProviders';
import {
    TYPE_TOKEN_LEGEND,
    VbaTypeSemanticTokensProvider,
} from './vbaSemanticTokensProvider';
import { registerVbaDiagnostics } from './vbaLiveDiagnostics';
import type { AnalysisWorkerClient } from './analysisWorkerClient';
import {
    VbaCodeActionProvider,
    XLIDE_ANALYZE_CURRENT_MODULE_ACTION_KIND,
    XLIDE_EXPORT_CURRENT_MODULE_ACTION_KIND,
} from './vbaCodeActions';
import {
    registerVbaAutoBlock,
    registerVbaLoopIteratorSync,
} from './vbaTypingAutomation';
import { VbaSymbolIndex } from './vbaSymbolIndex';
import { registerVbaMemberCompletion } from './vbaMemberCompletion';
import { DocMetadataLoader } from './vbaDocMetadata';
import { VbaProjectIndexService } from './vbaProjectIndexService';

const VBA_SELECTOR: vscode.DocumentSelector = [
    { scheme: XLIDE_SCHEME, language: 'vba' },
    { scheme: XLIDE_SCHEME, language: XLIDE_VBA_LANGUAGE_ID },
    { scheme: XLIDE_SCHEME },
    { language: 'vba' },
    { language: XLIDE_VBA_LANGUAGE_ID },
];

export function registerVbaLanguageProviders(
    context: vscode.ExtensionContext,
    bridge: PythonBridge,
    workerClient?: AnalysisWorkerClient,
): VbaSymbolIndex {
    const index = new VbaSymbolIndex(bridge);
    const projectIndexService = new VbaProjectIndexService(index);

    registerVbaDiagnostics(context, projectIndexService, workerClient);
    registerVbaAutoBlock(context);
    registerVbaLoopIteratorSync(context);
    const docMetadata = new DocMetadataLoader();
    void docMetadata.start(context);
    registerVbaMemberCompletion(context, projectIndexService, VBA_SELECTOR, docMetadata.registry);

    const typeSemanticTokensProvider = new VbaTypeSemanticTokensProvider(projectIndexService);
    context.subscriptions.push(
        index,
        projectIndexService,
        typeSemanticTokensProvider,
        vscode.languages.registerDocumentSymbolProvider(
            VBA_SELECTOR,
            new VbaDocumentSymbolProvider(),
            { label: 'XLIDE VBA' },
        ),
        vscode.languages.registerWorkspaceSymbolProvider(
            new VbaWorkspaceSymbolProvider(projectIndexService),
        ),
        vscode.languages.registerDefinitionProvider(
            VBA_SELECTOR,
            new VbaDefinitionProvider(projectIndexService),
        ),
        vscode.languages.registerReferenceProvider(
            VBA_SELECTOR,
            new VbaReferenceProvider(projectIndexService),
        ),
        vscode.languages.registerRenameProvider(
            VBA_SELECTOR,
            new VbaRenameProvider(projectIndexService),
        ),
        vscode.languages.registerCodeActionsProvider(
            VBA_SELECTOR,
            new VbaCodeActionProvider(),
            {
                providedCodeActionKinds: [
                    vscode.CodeActionKind.QuickFix,
                    XLIDE_ANALYZE_CURRENT_MODULE_ACTION_KIND,
                    XLIDE_EXPORT_CURRENT_MODULE_ACTION_KIND,
                ],
            },
        ),
        vscode.languages.registerDocumentSemanticTokensProvider(
            VBA_SELECTOR,
            typeSemanticTokensProvider,
            TYPE_TOKEN_LEGEND,
        ),
        // Keep the index consistent with saves to virtual VBA documents.
        vscode.workspace.onDidSaveTextDocument((doc) => {
            if (doc.uri.scheme !== XLIDE_SCHEME) { return; }
            try {
                const { xlsmPath, moduleName } = decodeModuleUri(doc.uri);
                index.updateModuleSource(xlsmPath, moduleName, doc.getText());
            } catch {
                // Ignore URIs we cannot decode.
            }
        }),
    );

    return index;
}
