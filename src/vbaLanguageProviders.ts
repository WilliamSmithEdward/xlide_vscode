import * as vscode from 'vscode';
import * as path from 'path';
import { PythonBridge } from './pythonBridge';
import {
    XLIDE_SCHEME,
    XLIDE_VBA_LANGUAGE_ID,
    decodeModuleUri,
    encodeModuleUri,
    isVbaDocument,
    moduleIdentityKey,
    workbookIdentityKey,
} from './xlideFileSystem';
import {
    isStandaloneVbaDocument,
    liveProjectIndexForDocument,
    moduleKindFromDocument,
    moduleNameFromDocument,
} from './vbaDocumentIdentity';
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
import {
    VbaCodeActionProvider,
    XLIDE_ANALYZE_CURRENT_MODULE_ACTION_KIND,
    XLIDE_EXPORT_CURRENT_MODULE_ACTION_KIND,
} from './vbaCodeActions';
import { VbaSymbolIndex, VbaModuleSymbols } from './vbaSymbolIndex';
import {
    findIdentifierOccurrences,
    lineStartOffsets,
    stripVba,
    VBA_IDENTIFIER_NAME_RE,
    VBA_IDENTIFIER_RE,
} from './vbaSourceScan';
import { analyzeVbaStructure } from './vbaStructuralDiagnostics';
import {
    detectSmartBlockOpener,
    isSmartBlockClosedAhead,
    procedureHeaderParensEdit,
    resolveLoopIteratorSyncEdit,
    smartBlockInsertion,
    withMemberContinuationText,
} from './vbaSmartEnter';
import {
    diagnosticSourceForCode,
    DiagnosticSeverity as RuleSeverity,
    EventHandlerDocumentType,
    eventHandlerDocumentTypeForContext,
    isXlideDiagnosticSource,
    normalizeDiagnosticCode,
    ProjectIndex,
    ReferenceScope,
    resolveDiagnosticCodeActions,
    resolveMemberDefinitionsAt,
    resolveProcedureLabelDefinitionAt,
    tokenizeCached,
    type MemberCompletionContext,
    resolveTypeReferenceAt,
    resolveTypeSemanticTokens,
    TypeSemanticTokenType,
    type VbaDiagnosticData,
    type VbaProjectClassMember,
    type VbaProjectClassMemberDefinition,
    type ModuleSymbolKind,
    type Span,
    VbaSymbol as AstSymbol,
    type VbaSymbolKind,
} from './analyzer';
import { analyzeVbaModuleSource } from './vbaModuleAnalysis';
import { registerVbaMemberCompletion } from './vbaMemberCompletion';
import { DocMetadataLoader } from './vbaDocMetadata';
import {
    createOffsetToPositionConverter,
    moduleKindFromType,
    offsetToPosition,
    projectTypeDefinitionToLocation,
    typeDefinitionsForReference,
    typeReferenceLocations,
} from './vbaNavigation';
import {
    buildLiveVbaProjectIndexAsync,
    projectAnalysisOptionsForModule,
    projectProcedureSignatures,
    type VbaProjectAnalysisOptions,
} from './vbaProjectAnalysis';
import { VbaProjectIndexService } from './vbaProjectIndexService';
import {
    documentOutlineSymbolsForSource,
    workspaceSymbols as presentedWorkspaceSymbols,
    type VbaPresentedSymbol,
    type VbaPresentedWorkspaceSymbol,
} from './vbaSymbolPresentation';
import {
    isAnalysisRuleTracked,
} from './analysisSettingsCore';
import {
    effectiveWorkbookAnalysisSettings,
    type EffectiveWorkbookAnalysisSettings,
} from './workbookAnalysisSettings';
import { startPerformanceTrace } from './performanceTrace';
import { isWorkbookSettingsError, settingsPathForWorkbook } from './workbookSettings';
import {
    validateXlideGlobalSettingsFromConfig,
    xlideDiagnosticsEnabledFromConfig,
    xlideEditorBlockLayoutFromConfig,
    type XlideGlobalSettingsProblem,
} from './globalSettings';
import { errorMessage } from './util/errors';
import {
    XLIDE_DIAGNOSTIC_DATA,
    type XlideDiagnosticWithData,
} from './xlideDiagnosticData';

const VBA_SELECTOR: vscode.DocumentSelector = [
    { scheme: XLIDE_SCHEME, language: 'vba' },
    { scheme: XLIDE_SCHEME, language: XLIDE_VBA_LANGUAGE_ID },
    { scheme: XLIDE_SCHEME },
    { language: 'vba' },
    { language: XLIDE_VBA_LANGUAGE_ID },
];
/**
 * VBA-IDE-style smart Enter: typing a block opener and pressing Enter
 * auto-inserts the matching closer below, leaving the cursor on the indented
 * body line. `With` also seeds the body line with `.` so member completion can
 * start immediately.
 */
function registerVbaAutoBlock(context: vscode.ExtensionContext): void {
    let applying = false;

    const sub = vscode.workspace.onDidChangeTextDocument(async (e) => {
        if (applying) { return; }
        const doc = e.document;
        if (!isVbaDocument(doc)) { return; }
        if (e.contentChanges.length !== 1) { return; }

        const change = e.contentChanges[0];
        // React only to a plain Enter (newline plus optional auto-indent),
        // never to pastes or multi-character insertions.
        if (!/^\r?\n[ \t]*$/.test(change.text)) { return; }

        const openerLineIndex = change.range.start.line;
        const openerLine = doc.lineAt(openerLineIndex).text;
        const headerParensEdit = procedureHeaderParensEdit(openerLine);
        const normalizedOpenerLine = headerParensEdit
            ? `${openerLine.slice(0, headerParensEdit.startCol)}${headerParensEdit.newText}${openerLine.slice(headerParensEdit.endCol)}`
            : openerLine;
        const opener = detectSmartBlockOpener(stripVba(normalizedOpenerLine));
        if (!opener) {
            await maybeContinueWithMemberLine(doc, openerLineIndex);
            return;
        }

        const bodyLineIndex = openerLineIndex + 1;
        if (bodyLineIndex >= doc.lineCount) { return; }

        const strippedLines = doc.getText().split(/\r\n|\r|\n/).map(stripVba);
        strippedLines[openerLineIndex] = stripVba(normalizedOpenerLine);
        const closedAhead = isSmartBlockClosedAhead(strippedLines, openerLineIndex, opener);

        const editor = vscode.window.activeTextEditor;
        if (!editor || editor.document !== doc) { return; }

        const eol = doc.eol === vscode.EndOfLine.CRLF ? '\r\n' : '\n';
        const bodyLine = doc.lineAt(bodyLineIndex).text;
        if (!/^[ \t]*$/.test(bodyLine)) { return; }
        const smartBlock = smartBlockInsertion(normalizedOpenerLine, bodyLine, opener, {
            eol,
            insertCloser: !closedAhead,
            layout: xlideEditorBlockLayoutFromConfig(vscode.workspace.getConfiguration('xlide')).value,
        });
        const bodyRange = new vscode.Range(
            new vscode.Position(bodyLineIndex, 0),
            new vscode.Position(bodyLineIndex, bodyLine.length),
        );

        applying = true;
        try {
            await editor.edit(
                (eb) => {
                    if (headerParensEdit) {
                        eb.insert(
                            new vscode.Position(openerLineIndex, headerParensEdit.startCol),
                            headerParensEdit.newText,
                        );
                    }
                    eb.replace(
                        bodyRange,
                        smartBlock.replacementText,
                    );
                },
                { undoStopBefore: false, undoStopAfter: true },
            );
        } finally {
            applying = false;
        }

        // Keep the caret on the indented body line, above the inserted End. The
        // delayed pass wins same-Enter listener races such as canonical casing.
        const placeCaret = (): void => {
            if (vscode.window.activeTextEditor !== editor || editor.document !== doc) {
                return;
            }
            const caretLineIndex = bodyLineIndex + smartBlock.bodyLineOffset;
            if (caretLineIndex >= doc.lineCount || doc.lineAt(caretLineIndex).text !== smartBlock.bodyText) {
                return;
            }
            const caret = new vscode.Position(
                caretLineIndex,
                smartBlock.bodyText.length,
            );
            editor.selection = new vscode.Selection(caret, caret);
        };
        placeCaret();
        setTimeout(placeCaret, 0);
    });

    context.subscriptions.push(sub);
}

async function maybeContinueWithMemberLine(
    doc: vscode.TextDocument,
    previousLineIndex: number,
): Promise<void> {
    const bodyLineIndex = previousLineIndex + 1;
    if (bodyLineIndex >= doc.lineCount) { return; }

    const bodyLine = doc.lineAt(bodyLineIndex).text;
    if (!/^[ \t]*$/.test(bodyLine)) { return; }

    const lineText = withMemberContinuationText(doc.getText(), previousLineIndex);
    if (!lineText) { return; }

    const editor = vscode.window.activeTextEditor;
    if (!editor || editor.document !== doc) { return; }

    const bodyRange = new vscode.Range(
        new vscode.Position(bodyLineIndex, 0),
        new vscode.Position(bodyLineIndex, bodyLine.length),
    );
    const applied = await editor.edit(
        (eb) => eb.replace(bodyRange, lineText),
        { undoStopBefore: false, undoStopAfter: true },
    );
    if (!applied) { return; }

    const placeCaret = (): void => {
        if (vscode.window.activeTextEditor !== editor || editor.document !== doc) {
            return;
        }
        if (bodyLineIndex >= doc.lineCount || doc.lineAt(bodyLineIndex).text !== lineText) {
            return;
        }
        const caret = new vscode.Position(bodyLineIndex, lineText.length);
        editor.selection = new vscode.Selection(caret, caret);
    };
    placeCaret();
    setTimeout(placeCaret, 0);
}

/**
 * Keeps simple loop iterator names paired across `For`/`For Each` and `Next`.
 * This intentionally lives outside snippets so hand-written loops get the same
 * behavior as completed loops.
 */
function registerVbaLoopIteratorSync(context: vscode.ExtensionContext): void {
    let applying = false;

    const sub = vscode.workspace.onDidChangeTextDocument(async (e) => {
        if (applying) { return; }
        const doc = e.document;
        if (!isVbaDocument(doc)) { return; }
        if (e.contentChanges.length !== 1) { return; }

        const change = e.contentChanges[0];
        if (/[\r\n]/.test(change.text)) { return; }

        const editor = vscode.window.activeTextEditor;
        if (!editor || editor.document !== doc) { return; }

        const lineIndex = Math.min(change.range.start.line, doc.lineCount - 1);
        const lineLength = doc.lineAt(lineIndex).text.length;
        const character = Math.min(lineLength, change.range.start.character + change.text.length);
        const offset = doc.offsetAt(new vscode.Position(lineIndex, character));
        const syncEdit = resolveLoopIteratorSyncEdit(doc.getText(), offset);
        if (!syncEdit) { return; }

        applying = true;
        try {
            await editor.edit(
                (eb) => eb.replace(
                    new vscode.Range(
                        doc.positionAt(syncEdit.span.start),
                        doc.positionAt(syncEdit.span.end),
                    ),
                    syncEdit.newText,
                ),
                { undoStopBefore: false, undoStopAfter: false },
            );
        } finally {
            applying = false;
        }
    });

    context.subscriptions.push(sub);
}

export function registerVbaLanguageProviders(
    context: vscode.ExtensionContext,
    bridge: PythonBridge,
): VbaSymbolIndex {
    const index = new VbaSymbolIndex(bridge);
    const projectIndexService = new VbaProjectIndexService(index);

    registerVbaDiagnostics(context, projectIndexService);
    registerVbaAutoBlock(context);
    registerVbaLoopIteratorSync(context);
    const docMetadata = new DocMetadataLoader();
    void docMetadata.start(context);
    registerVbaMemberCompletion(context, projectIndexService, VBA_SELECTOR, docMetadata.registry);

    context.subscriptions.push(
        index,
        projectIndexService,
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
            new VbaTypeSemanticTokensProvider(projectIndexService),
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
