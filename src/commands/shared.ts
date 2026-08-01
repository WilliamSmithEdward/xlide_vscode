import * as vscode from 'vscode';
import type { WorkbookEngine } from '../workbookEngine';
import type { XlsmExplorer, XlideNode } from '../xlsmExplorer';
import type { XlideFileSystemProvider } from '../xlideFileSystem';
import type { VbaSymbolIndex } from '../vbaSymbolIndex';
import {
    XLIDE_SCHEME,
    activeLocalVbaEditor,
    decodeModuleUri,
} from '../xlideFileSystem';
import {
    formatChangeSummaryDetails,
    type XlideChangeSummary,
} from '../xlideWriteAudit';
import { parseModule } from '../analyzer/parser/parseModule';
import type { ModuleMember } from '../analyzer/parser/nodes';
import { spanContainsOffset, spanLength } from '../vbaAnalysisSuppression';

/**
 * The activation-time singletons every command domain registers against.
 * Passed explicitly (instead of captured in one giant closure) so each
 * per-domain `registerXxxCommands(deps)` module stays independently readable.
 */
export interface CommandDeps {
    context: vscode.ExtensionContext;
    bridge: WorkbookEngine;
    explorer: XlsmExplorer;
    fsProvider: XlideFileSystemProvider;
    out: vscode.OutputChannel;
    vbaIndex: VbaSymbolIndex;
}

/**
 * Transient status-bar confirmation for successful command results whose
 * outcome is already visible elsewhere (an opened results panel, an editor
 * edit, files on disk). Replaces popup toasts, which interrupt without adding
 * information; failures and actionable prompts stay as notifications.
 */
export function statusMessage(text: string): void {
    vscode.window.setStatusBarMessage(text, 6000);
}

/** Logs every detail line of a change summary and returns the headline line. */
export function logChangeSummary(
    log: (msg: string) => void,
    prefix: string,
    summary: XlideChangeSummary,
): string {
    const lines = formatChangeSummaryDetails(summary);
    for (const line of lines) {
        log(`[${prefix}] ${line}`);
    }
    return lines[0];
}

/** Workbook path from an explorer node, falling back to the active XLIDE editor. */
export function resolveWorkbookPath(node?: XlideNode): string | undefined {
    let filePath = node?.filePath;
    if (!filePath) {
        const active = vscode.window.activeTextEditor;
        if (active && active.document.uri.scheme === XLIDE_SCHEME) {
            filePath = decodeModuleUri(active.document.uri).xlsmPath;
        }
    }
    return filePath;
}

export async function activeLocalWorkbookPath(): Promise<string | undefined> {
    const editor = activeLocalVbaEditor();
    return editor ? decodeModuleUri(editor.document.uri).xlsmPath : undefined;
}

/**
 * Shows a document, retrying briefly when VS Code surfaces a different editor
 * (e.g. while a webview panel steals focus during analysis navigation).
 */
export async function showAnalysisSourceDocument(
    doc: vscode.TextDocument,
    viewColumn?: vscode.ViewColumn,
): Promise<vscode.TextEditor> {
    let lastError: unknown;
    let lastEditor: vscode.TextEditor | undefined;
    for (let attempt = 0; attempt < 3; attempt++) {
        if (attempt > 0) {
            await delay(50 * attempt);
        }
        try {
            const editor = await vscode.window.showTextDocument(doc, { preview: false, viewColumn });
            if (sameDocumentUri(editor.document.uri, doc.uri)) {
                return editor;
            }
            lastEditor = editor;
        } catch (err) {
            lastError = err;
        }
    }
    if (lastEditor) {
        return lastEditor;
    }
    throw lastError instanceof Error
        ? lastError
        : new Error('VS Code did not open the analysis source document.');
}

function sameDocumentUri(left: vscode.Uri, right: vscode.Uri): boolean {
    return left.toString() === right.toString();
}

function delay(milliseconds: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, milliseconds));
}

type ProcedureMember = Extract<ModuleMember, { kind: 'Procedure' }>;

/** Innermost procedure containing the cursor, if any. */
export function procedureAtCursor(editor: vscode.TextEditor): ProcedureMember | undefined {
    const source = editor.document.getText();
    const offset = editor.document.offsetAt(editor.selection.active);
    const parsed = parseModule(source);
    return parsed.members
        .filter((member): member is ProcedureMember => member.kind === 'Procedure')
        .filter((member) => spanContainsOffset(member.span, offset))
        .sort((left, right) => spanLength(left.span) - spanLength(right.span))[0];
}

/** Name of the innermost procedure containing the cursor, if any. */
export function procedureNameAtCursor(editor: vscode.TextEditor): string | undefined {
    return procedureAtCursor(editor)?.name;
}

/**
 * Names of the parameters a caller MUST supply (not Optional, not ParamArray).
 * A procedure with any of these cannot be run by F5, which passes no arguments.
 */
export function requiredParameterNames(procedure: ProcedureMember): string[] {
    return procedure.params
        .filter((param) => !param.optional && !param.paramArray)
        .map((param) => param.name);
}
