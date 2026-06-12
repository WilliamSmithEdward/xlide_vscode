// VBA code actions: quick fixes resolved from analyzer diagnostics (via the
// shared diagnostic-data symbol) plus the XLIDE source actions for analyzing
// and exporting the current module.
//
// Extracted verbatim from vbaLanguageProviders.ts (audit #21).

import * as vscode from 'vscode';
import { XLIDE_SCHEME, isVbaDocument } from './xlideFileSystem';
import { analyzeVbaStructure } from './vbaStructuralDiagnostics';
import {
    isXlideDiagnosticSource,
    normalizeDiagnosticCode,
    resolveDiagnosticCodeActions,
} from './analyzer';
import {
    XLIDE_DIAGNOSTIC_DATA,
    type XlideDiagnosticWithData,
} from './xlideDiagnosticData';

const XLIDE_SOURCE_ACTION_KIND = vscode.CodeActionKind.Source.append('xlide');
export const XLIDE_ANALYZE_CURRENT_MODULE_ACTION_KIND = XLIDE_SOURCE_ACTION_KIND.append('analyzeCurrentModule');
export const XLIDE_EXPORT_CURRENT_MODULE_ACTION_KIND = XLIDE_SOURCE_ACTION_KIND.append('exportCurrentModule');

export class VbaCodeActionProvider implements vscode.CodeActionProvider {
    public provideCodeActions(
        document: vscode.TextDocument,
        _range: vscode.Range,
        context: vscode.CodeActionContext,
        _token: vscode.CancellationToken,
    ): vscode.ProviderResult<vscode.CodeAction[]> {
        if (!isVbaDocument(document)) { return []; }
        const wantsQuickFix = codeActionKindRequested(context.only, vscode.CodeActionKind.QuickFix);
        const wantsAnalyzeCurrentModule = codeActionKindRequested(context.only, XLIDE_ANALYZE_CURRENT_MODULE_ACTION_KIND);
        const wantsExportCurrentModule = codeActionKindRequested(context.only, XLIDE_EXPORT_CURRENT_MODULE_ACTION_KIND);
        if (!wantsQuickFix && !wantsAnalyzeCurrentModule && !wantsExportCurrentModule) {
            return [];
        }

        const source = document.getText();
        const actions: vscode.CodeAction[] = [];
        if (wantsAnalyzeCurrentModule && document.uri.scheme === XLIDE_SCHEME) {
            const action = new vscode.CodeAction(
                'XLIDE: Analyze Current Module',
                XLIDE_ANALYZE_CURRENT_MODULE_ACTION_KIND,
            );
            action.command = {
                command: 'xlide.analyzeCurrentModule',
                title: 'Analyze Current Module',
            };
            actions.push(action);
        }
        if (wantsExportCurrentModule && document.uri.scheme === XLIDE_SCHEME && !document.uri.authority) {
            const action = new vscode.CodeAction(
                'XLIDE: Export/Sync Current Module',
                XLIDE_EXPORT_CURRENT_MODULE_ACTION_KIND,
            );
            action.command = {
                command: 'xlide.exportCurrentModuleToFolder',
                title: 'Export/Sync Current Module',
            };
            actions.push(action);
        }
        if (!wantsQuickFix) {
            return actions;
        }
        let structuralDiagnostics: ReturnType<typeof analyzeVbaStructure> | undefined;
        for (const diagnostic of context.diagnostics) {
            if (!isXlideDiagnosticSource(diagnostic.source)) { continue; }
            const code = normalizeDiagnosticCode(diagnostic.code);
            if (!code) { continue; }
            const structuralDiagnostic = code === 'missing-block-closer'
                ? matchingStructuralDiagnostic(
                    diagnostic,
                    structuralDiagnostics ??= analyzeVbaStructure(source),
                )
                : undefined;
            const fixes = resolveDiagnosticCodeActions(source, {
                code,
                message: diagnostic.message,
                expectedClose: structuralDiagnostic?.expectedClose,
                insertLine: structuralDiagnostic?.insertLine,
                expectedCloseReplacementSpan: structuralDiagnostic?.expectedCloseReplacement
                    ? {
                        start: document.offsetAt(new vscode.Position(
                            structuralDiagnostic.expectedCloseReplacement.line,
                            structuralDiagnostic.expectedCloseReplacement.startCol,
                        )),
                        end: document.offsetAt(new vscode.Position(
                            structuralDiagnostic.expectedCloseReplacement.line,
                            structuralDiagnostic.expectedCloseReplacement.endCol,
                        )),
                    }
                    : undefined,
                expectedCloseReplacementText: structuralDiagnostic?.expectedCloseReplacement?.text,
                span: {
                    start: document.offsetAt(diagnostic.range.start),
                    end: document.offsetAt(diagnostic.range.end),
                },
                includeSuppressionAction: true,
                data: (diagnostic as XlideDiagnosticWithData)[XLIDE_DIAGNOSTIC_DATA],
            });
            for (const fix of fixes) {
                const action = new vscode.CodeAction(fix.title, vscode.CodeActionKind.QuickFix);
                action.diagnostics = [diagnostic];
                action.isPreferred = fix.isPreferred;
                const edit = new vscode.WorkspaceEdit();
                for (const textEdit of fix.edits) {
                    edit.replace(
                        document.uri,
                        new vscode.Range(
                            document.positionAt(textEdit.span.start),
                            document.positionAt(textEdit.span.end),
                        ),
                        textEdit.newText,
                    );
                }
                action.edit = edit;
                actions.push(action);
            }
        }
        return actions;
    }
}

function codeActionKindRequested(
    only: vscode.CodeActionKind | undefined,
    kind: vscode.CodeActionKind,
): boolean {
    return !only || only.contains(kind) || kind.contains(only);
}

function matchingStructuralDiagnostic(
    diagnostic: vscode.Diagnostic,
    problems: ReturnType<typeof analyzeVbaStructure>,
): ReturnType<typeof analyzeVbaStructure>[number] | undefined {
    return problems.find((problem) =>
        problem.code === normalizeDiagnosticCode(diagnostic.code) &&
        problem.message === diagnostic.message &&
        diagnostic.range.start.isEqual(new vscode.Position(problem.line, problem.startCol)) &&
        diagnostic.range.end.isEqual(new vscode.Position(problem.line, problem.endCol)),
    );
}
