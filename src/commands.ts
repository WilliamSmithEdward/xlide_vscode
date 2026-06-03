import * as vscode from 'vscode';
import * as path from 'path';
import * as cp from 'child_process';
import * as fs from 'fs';
import * as os from 'os';
import { PythonBridge } from './pythonBridge';
import { XlsmExplorer, XlideNode } from './xlsmExplorer';
import {
    XlideFileSystemProvider,
    encodeModuleUri,
    decodeModuleUri,
    XLIDE_SCHEME,
    notifySignatureDropped,
} from './xlideFileSystem';
import { applyOpenDocumentSources } from './vbaOpenDocuments';
import { encodeRemoteModuleUri } from './liveShare';
import {
    exportWorkbookModule,
} from './moduleExport';
import {
    type ExportMode,
} from './workbookSettings';
import {
    analyzeWorkbook,
    summarizeWorkbookAnalysisProblems,
    workbookProblemsForModule,
    type WorkbookAnalysisProblem,
    type WorkbookAnalysisResult,
} from './vbaWorkbookAnalysis';
import {
    openWorkbookAnalysisResults,
    type WorkbookAnalysisSuppressScope,
} from './workbookAnalysisWebview';
import {
    createVbaTestRunReport,
    describeVbaTestSelection,
    discoverWorkbookVbaTests,
    summarizeVbaTestTags,
    summarizeVbaTestRun,
    vbaTestFailureMessage,
    type VbaTestCase,
    type VbaTestRunItem,
    type VbaTestRunReport,
    type VbaTestSelectionOptions,
} from './vbaTestRunner';
import {
    buildVbaTestDirectRunnerModule,
    buildOwnedReadOnlyExcelTestHostScript,
    DEFAULT_VBA_TEST_TIMEOUT_MS,
    parseVbaTestHostEventLine,
    vbaTestHostPlanItems,
} from './vbaTestExcelHost';
import {
    validateVbaTestHostOracleTrace,
    type VbaTestHostOracleEvent,
} from './vbaTestHostOracle';
import { writeVbaTestRunArtifacts } from './vbaTestArtifacts';
import { openVbaTestResults, setVbaTestResultsRunning } from './vbaTestResultsWebview';
import { checkExcelComAvailability } from './excelComAvailability';
import {
    openVbaTestsPanel,
    type VbaTestsRunFilterRequest,
    type VbaTestSupportStatusModel,
    type VbaTestsPanelModel,
} from './vbaTestsWebview';
import {
    normalizeVbaTestSupportModuleSource,
    XLIDE_ASSERT_MODULE_NAME,
    XLIDE_ASSERT_MODULE_SOURCE,
} from './vbaTestSupportModule';
import { invalidateVbaMemberCompletionCache } from './vbaMemberCompletion';
import { analyzeVbaModuleSource } from './vbaModuleAnalysis';
import {
    resolvedXlideGlobalSettingsFromConfig,
    xlideAttachToRunningExcelFromConfig,
} from './globalSettings';
import { effectiveWorkbookAnalysisSettings } from './workbookAnalysisSettings';
import { effectiveWorkbookTestSettings } from './workbookTestSettings';
import { lineStartOffsets, VBA_IDENTIFIER_NAME_RE } from './vbaStructuralAnalysis';
import { VbaSymbolIndex } from './vbaSymbolIndex';
import {
    buildVbaProjectIndex,
    moduleKindFromType,
    projectClassModuleDefinition,
} from './vbaNavigation';
import {
    buildLiveVbaProjectIndex,
    projectAnalysisOptionsForModule,
    projectProcedureSignatures,
} from './vbaProjectAnalysis';
import {
    projectClassReferenceEdit,
    renameProjectClassModule,
} from './vbaClassRename';
import { resolveDiagnosticCodeActions } from './analyzer';
import {
    anonymizedWorkbookAnalysisReportFromResult,
    buildSupportBundle,
    defaultSupportBundleFileName,
    supportBundleDisclosureText,
    supportDiagnosticsText,
    type SupportBundle,
    type SupportBundleAnonymizedAnalysisReport,
    type SupportBundleAnalysisSummary,
    type SupportBundleSetting,
    type SupportBundleWorkbookSummary,
} from './supportBundle';
import {
    errorCategoryForSupportLog,
    recentXlideCommands,
} from './xlideCommandLog';
import { registerXlideCommand } from './xlideCommandRegistration';
import { recentXlideOutputLog } from './xlideOutputLog';
import {
    formatChangeSummary,
    formatChangeSummaryDetails,
    recentXlideWriteAudits,
    recordXlideWriteAudit,
    type XlideChangeSummary,
    type XlideWriteAuditOutcome,
} from './xlideWriteAudit';
import {
    buildExportModuleSyncPlan,
    buildImportModuleSyncPlan,
    type ImportMode,
    type ModuleSyncPlan,
    type ModuleSyncPlanItem,
} from './moduleSyncPlan';
import {
    openModuleSyncPreview,
    type ModuleSyncApplyResult,
    type ModuleSyncSettings,
} from './moduleSyncWebview';
import {
    effectiveWorkbookModuleSyncSettings,
    updateWorkbookModuleSyncSettings,
    type WorkbookModuleSyncFolderSource,
    type WorkbookModuleSyncModeSource,
} from './workbookModuleSyncSettings';
import { parseModule } from './analyzer/parser/parseModule';
import type { BodyNode, ModuleMember, Span } from './analyzer/parser/nodes';

type AnalysisSuppressionInsertionTarget =
    | { kind: 'module'; startLine: number }
    | { kind: 'member'; startLine: number }
    | { kind: 'block'; startLine: number; endLine: number };

type SuppressibleMember = Extract<ModuleMember, { kind: 'Procedure' | 'Type' | 'Enum' }>;
type BlockBodyNode = Extract<BodyNode, { body: BodyNode[] }>;

interface ResolvedModuleSyncSettings extends ModuleSyncSettings {
    folderPathSource: WorkbookModuleSyncFolderSource;
    exportModeSource?: WorkbookModuleSyncModeSource;
    importModeSource?: WorkbookModuleSyncModeSource;
    settingsPath: string;
}

interface VbaTestRunOptions {
    selection?: VbaTestSelectionOptions;
    failFast?: boolean;
}

interface VbaTestLastFailedRun {
    testIds: string[];
    tests: Array<{
        id: string;
        qualifiedName: string;
        status: VbaTestRunItem['status'];
    }>;
}

interface OwnedReadOnlyExcelHostRunResult {
    events: VbaTestHostOracleEvent[];
    resultsByName: Map<string, OwnedReadOnlyExcelHostTestResult>;
    hostError?: string;
    timedOutAfter?: string;
}

interface VbaTestRunExecution {
    report: VbaTestRunReport;
    hostEvents: VbaTestHostOracleEvent[];
}

interface OwnedReadOnlyExcelHostTestResult {
    outcome: 'passed' | 'failed' | 'timeout' | 'modal-blocked' | 'runner-error';
    durationMs: number;
    message?: string;
}

type OwnedExcelKillReason = 'timeout' | 'hung' | 'modal-blocked' | 'runner-error' | 'cleanup-failed';
const DEFAULT_VBA_TEST_CLEANUP_GRACE_MS = 5000;
const VBA_TEST_RERUN_FAILED_STATUSES = new Set<VbaTestRunItem['status']>([
    'failed',
    'timeout',
    'host-error',
    'xpass',
]);

type ProcedureMember = Extract<ModuleMember, { kind: 'Procedure' }>;

function psSingleQuoted(value: string): string {
    return `'${value.replace(/'/g, "''")}'`;
}

function suppressionTargetForProblem(
    source: string,
    starts: readonly number[],
    problemOffset: number,
    scope: WorkbookAnalysisSuppressScope,
): AnalysisSuppressionInsertionTarget {
    if (scope === 'module') {
        return { kind: 'module', startLine: moduleSuppressionInsertLine(source) };
    }

    const parsed = parseModule(source);
    const member = containingSuppressibleMember(parsed.members, problemOffset);
    if (!member) {
        throw new Error(`No containing Sub, Function, Property, Type, or Enum was found for this analysis finding.`);
    }

    if (scope === 'member') {
        return {
            kind: 'member',
            startLine: lineForOffset(starts, member.span.start),
        };
    }

    if (member.kind !== 'Procedure') {
        throw new Error('No containing executable block was found for this analysis finding.');
    }
    const block = closestContainingBlock(member.body, problemOffset);
    if (!block) {
        throw new Error('No containing executable block was found for this analysis finding.');
    }

    return {
        kind: 'block',
        startLine: lineForOffset(starts, block.span.start),
        endLine: lineForOffset(starts, Math.max(block.span.start, block.span.end - 1)),
    };
}

function moduleSuppressionInsertLine(source: string): number {
    const lines = source.split(/\r\n|\r|\n/);
    let line = 0;
    while (line < lines.length && /^\s*Attribute\b/i.test(lines[line])) {
        line++;
    }
    return line;
}

function containingSuppressibleMember(
    members: readonly ModuleMember[],
    offset: number,
): SuppressibleMember | undefined {
    return members
        .filter((member): member is SuppressibleMember =>
            member.kind === 'Procedure' || member.kind === 'Type' || member.kind === 'Enum',
        )
        .filter((member) => spanContainsOffset(member.span, offset))
        .sort((left, right) => spanLength(left.span) - spanLength(right.span))[0];
}

function closestContainingBlock(nodes: readonly BodyNode[], offset: number): BlockBodyNode | undefined {
    let best: BlockBodyNode | undefined;
    for (const node of nodes) {
        if (!isBlockBodyNode(node) || !spanContainsOffset(node.span, offset)) {
            continue;
        }
        const nested = closestContainingBlock(node.body, offset);
        const candidate = nested ?? node;
        if (!best || spanLength(candidate.span) < spanLength(best.span)) {
            best = candidate;
        }
    }
    return best;
}

function isBlockBodyNode(node: BodyNode): node is BlockBodyNode {
    return node.kind === 'IfBlock' ||
        node.kind === 'ForBlock' ||
        node.kind === 'DoBlock' ||
        node.kind === 'WhileBlock' ||
        node.kind === 'WithBlock' ||
        node.kind === 'SelectBlock';
}

function spanContainsOffset(span: Span, offset: number): boolean {
    return offset >= span.start && offset < Math.max(span.end, span.start + 1);
}

function spanLength(span: Span): number {
    return Math.max(1, span.end - span.start);
}

function lineForOffset(starts: readonly number[], offset: number): number {
    let lo = 0;
    let hi = starts.length - 1;
    while (lo < hi) {
        const mid = (lo + hi + 1) >> 1;
        if (starts[mid] <= offset) {
            lo = mid;
        } else {
            hi = mid - 1;
        }
    }
    return lo;
}

function copilotAnalysisPrompt(
    filePath: string,
    problem: WorkbookAnalysisProblem,
    source: string,
): string {
    const lines = source.split(/\r\n|\r|\n/);
    const zeroBasedLine = Math.max(0, problem.line - 1);
    const start = Math.max(0, zeroBasedLine - 6);
    const end = Math.min(lines.length, zeroBasedLine + 7);
    const excerpt = lines
        .slice(start, end)
        .map((line, index) => `${String(start + index + 1).padStart(4, ' ')}: ${line}`)
        .join('\n');
    const rule = problem.code
        ? `${problem.code}${problem.ruleTitle ? ` (${problem.ruleTitle})` : ''}`
        : problem.ruleTitle ?? 'unknown rule';
    return [
        'Please help me understand and fix this Excel VBA analysis finding from XLIDE.',
        '',
        `Workbook: ${path.basename(filePath)}`,
        `Module: ${problem.moduleName} (${problem.moduleType})`,
        `Location: ${problem.line}:${problem.column}`,
        `Severity: ${problem.severity}`,
        `Rule: ${rule}`,
        `Evidence: ${problem.diagnosticKind ?? 'unknown'}`,
        `Message: ${problem.message}`,
        '',
        'Relevant VBA source:',
        '```vba',
        excerpt,
        '```',
        '',
        'Please explain whether the finding is valid, what VBA rule or behavior is involved, and the smallest code change that would fix it.',
    ].join('\n');
}

export function registerCommands(
    context: vscode.ExtensionContext,
    bridge: PythonBridge,
    explorer: XlsmExplorer,
    _fsProvider: XlideFileSystemProvider,
    out: vscode.OutputChannel,
    vbaIndex: VbaSymbolIndex,
): vscode.Disposable[] {
    const lastFailedVbaTestRuns = new Map<string, VbaTestLastFailedRun>();

    function log(msg: string): void {
        out.appendLine(msg);
    }

    function logChangeSummary(prefix: string, summary: XlideChangeSummary): string {
        const lines = formatChangeSummaryDetails(summary);
        for (const line of lines) {
            log(`[${prefix}] ${line}`);
        }
        return lines[0];
    }

    function recordWriteAudit(input: {
        command: string;
        operation: string;
        outcome: XlideWriteAuditOutcome;
        workbookPath?: string;
        moduleName?: string;
        sourcePath?: string;
        targetPath?: string;
        summary: string;
        error?: unknown;
    }): void {
        recordXlideWriteAudit({
            timestamp: new Date().toISOString(),
            command: input.command,
            operation: input.operation,
            outcome: input.outcome,
            workbookPath: input.workbookPath,
            moduleName: input.moduleName,
            sourcePath: input.sourcePath,
            targetPath: input.targetPath,
            summary: input.summary,
            errorCategory: input.error ? errorCategoryForSupportLog(input.error) : undefined,
        });
    }

    async function openWorkbookAnalysisProblem(
        filePath: string,
        problem: WorkbookAnalysisProblem,
        analysisPanelColumn?: vscode.ViewColumn,
    ): Promise<void> {
        const uri = encodeModuleUri(filePath, problem.moduleName);
        const doc = await vscode.workspace.openTextDocument(uri);
        await vscode.languages.setTextDocumentLanguage(doc, 'vba');
        const editor = await vscode.window.showTextDocument(doc, {
            preview: false,
            viewColumn: adjacentAnalysisSourceColumn(analysisPanelColumn),
        });
        const line = Math.max(0, problem.line - 1);
        const startColumn = Math.max(0, problem.column - 1);
        const endColumn = Math.max(startColumn + 1, problem.endColumn - 1);
        const start = new vscode.Position(line, startColumn);
        const end = new vscode.Position(line, endColumn);
        const range = new vscode.Range(start, end);
        editor.selection = new vscode.Selection(start, end);
        editor.revealRange(range, vscode.TextEditorRevealType.InCenterIfOutsideViewport);
    }

    async function suppressWorkbookAnalysisProblem(
        filePath: string,
        problem: WorkbookAnalysisProblem,
        scope: WorkbookAnalysisSuppressScope,
        analysisPanelColumn?: vscode.ViewColumn,
    ): Promise<void> {
        const uri = encodeModuleUri(filePath, problem.moduleName);
        const doc = await vscode.workspace.openTextDocument(uri);
        await vscode.languages.setTextDocumentLanguage(doc, 'vba');
        const source = doc.getText();
        const starts = lineStartOffsets(source);
        const problemOffset = Math.max(
            0,
            (starts[Math.max(0, problem.line - 1)] ?? 0) + Math.max(0, problem.column - 1),
        );
        if (!problem.suppressionScopes.includes(scope)) {
            throw new Error(`Ignore ${scope} is not valid for '${problem.code ?? 'this analysis finding'}'.`);
        }
        const target = suppressionTargetForProblem(source, starts, problemOffset, scope);
        const code = (problem.code ?? 'all').trim() || 'all';
        const eol = doc.eol === vscode.EndOfLine.CRLF ? '\r\n' : '\n';
        const edit = new vscode.WorkspaceEdit();

        if (target.kind === 'block') {
            edit.insert(uri, new vscode.Position(target.endLine + 1, 0), `' @xlide-analysis-enable-block ${code}${eol}`);
            edit.insert(uri, new vscode.Position(target.startLine, 0), `' @xlide-analysis-disable-block ${code}${eol}`);
        } else if (target.kind === 'member') {
            edit.insert(uri, new vscode.Position(target.startLine, 0), `' @xlide-analysis-disable-next-member ${code}${eol}`);
        } else {
            edit.insert(uri, new vscode.Position(target.startLine, 0), `' @xlide-analysis-disable-file ${code}${eol}`);
        }

        const applied = await vscode.workspace.applyEdit(edit);
        if (!applied) {
            throw new Error('VS Code rejected the analysis ignore edit.');
        }
        const editor = await vscode.window.showTextDocument(doc, {
            preview: false,
            viewColumn: adjacentAnalysisSourceColumn(analysisPanelColumn),
        });
        const position = new vscode.Position(target.startLine, 0);
        editor.selection = new vscode.Selection(position, position);
        editor.revealRange(new vscode.Range(position, position), vscode.TextEditorRevealType.InCenterIfOutsideViewport);
        vscode.window.showInformationMessage(`XLIDE: Added ${scope} analysis ignore directive for '${code}'.`);
    }

    async function askCopilotAboutWorkbookAnalysisProblem(
        filePath: string,
        problem: WorkbookAnalysisProblem,
        analysisPanelColumn?: vscode.ViewColumn,
    ): Promise<void> {
        const uri = encodeModuleUri(filePath, problem.moduleName);
        const doc = await vscode.workspace.openTextDocument(uri);
        await vscode.languages.setTextDocumentLanguage(doc, 'vba');
        await openWorkbookAnalysisProblem(filePath, problem, analysisPanelColumn);
        const prompt = copilotAnalysisPrompt(filePath, problem, doc.getText());
        try {
            await vscode.commands.executeCommand('workbench.action.chat.open', { query: prompt });
        } catch {
            await vscode.env.clipboard.writeText(prompt);
            vscode.window.showWarningMessage(
                'XLIDE: Could not open VS Code Chat. The Copilot prompt was copied to the clipboard.',
            );
        }
    }

    async function quickFixWorkbookAnalysisProblem(
        filePath: string,
        problem: WorkbookAnalysisProblem,
        analysisPanelColumn?: vscode.ViewColumn,
        fixIndex = 0,
    ): Promise<boolean> {
        if (!problem.code) {
            return false;
        }

        const uri = encodeModuleUri(filePath, problem.moduleName);
        const doc = await vscode.workspace.openTextDocument(uri);
        await vscode.languages.setTextDocumentLanguage(doc, 'vba');
        const source = doc.getText();
        const starts = lineStartOffsets(source);
        const lineStart = starts[Math.max(0, problem.line - 1)] ?? 0;
        const span = {
            start: lineStart + Math.max(0, problem.column - 1),
            end: lineStart + Math.max(0, problem.endColumn - 1),
        };
        const fixes = resolveDiagnosticCodeActions(source, {
            code: problem.code,
            message: problem.message,
            span,
            expectedClose: problem.expectedClose,
            insertLine: problem.insertLine,
            data: problem.data,
            includeSuppressionAction: false,
        });
        const fix = fixes[fixIndex] ?? fixes[0];
        if (!fix || fix.edits.length === 0) {
            return false;
        }

        const edit = new vscode.WorkspaceEdit();
        for (const textEdit of fix.edits) {
            edit.replace(
                uri,
                new vscode.Range(
                    doc.positionAt(textEdit.span.start),
                    doc.positionAt(textEdit.span.end),
                ),
                textEdit.newText,
            );
        }
        const applied = await vscode.workspace.applyEdit(edit);
        if (!applied) {
            return false;
        }

        const editor = await vscode.window.showTextDocument(doc, {
            preview: false,
            viewColumn: adjacentAnalysisSourceColumn(analysisPanelColumn),
        });
        const firstEdit = fix.edits[0];
        const position = doc.positionAt(firstEdit.span.start);
        editor.selection = new vscode.Selection(position, position);
        editor.revealRange(new vscode.Range(position, position), vscode.TextEditorRevealType.InCenterIfOutsideViewport);
        return true;
    }

    function showWorkbookAnalysisResults(
        result: WorkbookAnalysisResult,
        onRefreshResult?: () => Promise<WorkbookAnalysisResult>,
    ): void {
        const filePath = result.filePath;
        openWorkbookAnalysisResults(context, result, {
            onOpenProblem: (problem, analysisPanelColumn) =>
                openWorkbookAnalysisProblem(filePath, problem, analysisPanelColumn),
            onQuickFixProblem: (problem, analysisPanelColumn, fixIndex) =>
                quickFixWorkbookAnalysisProblem(filePath, problem, analysisPanelColumn, fixIndex),
            onSuppressProblem: (problem, scope, analysisPanelColumn) =>
                suppressWorkbookAnalysisProblem(filePath, problem, scope, analysisPanelColumn),
            onAskCopilot: (problem, analysisPanelColumn) =>
                askCopilotAboutWorkbookAnalysisProblem(filePath, problem, analysisPanelColumn),
            onRefreshResult,
            onDidChangeWorkbookTree: explorer.onDidChangeTreeData,
        });
    }

    function adjacentAnalysisSourceColumn(analysisPanelColumn?: vscode.ViewColumn): vscode.ViewColumn {
        switch (analysisPanelColumn) {
            case vscode.ViewColumn.One:
                return vscode.ViewColumn.Two;
            case vscode.ViewColumn.Two:
            case vscode.ViewColumn.Three:
            case vscode.ViewColumn.Four:
            case vscode.ViewColumn.Five:
            case vscode.ViewColumn.Six:
            case vscode.ViewColumn.Seven:
            case vscode.ViewColumn.Eight:
            case vscode.ViewColumn.Nine:
                return vscode.ViewColumn.One;
            default:
                return vscode.ViewColumn.Beside;
        }
    }

    function shouldAttachToRunningExcel(): boolean {
        return xlideAttachToRunningExcelFromConfig(vscode.workspace.getConfiguration('xlide')).value;
    }

    function showRunMacroFailure(err: unknown): void {
        const raw = err instanceof Error ? err.message : String(err);
        const pipe = raw.indexOf('|');
        const code = pipe >= 0 ? raw.slice(0, pipe) : '';
        const message = pipe >= 0 ? raw.slice(pipe + 1) : raw;
        if (code === 'REOPEN_BLOCKED' || code === 'REOPEN_FAILED') {
            void vscode.window.showWarningMessage(`XLIDE: ${message}`);
            return;
        }
        void vscode.window.showErrorMessage(`XLIDE: Failed to run macro: ${message}`);
    }

    // Helper functions for Windows COM-based Excel operations
    function runWindowsExcel(filePath: string, attachToRunning: boolean, readOnly: boolean): void {
        const roFlag = readOnly ? '$true' : '$false';
        const script = [
            '$ErrorActionPreference = "Stop"',
            `$targetPath = ${psSingleQuoted(filePath)}`,
            `$targetName = ${psSingleQuoted(path.basename(filePath))}`,
            '$excel = $null',
            '$workbook = $null',
            `$attachToRunning = ${attachToRunning ? '$true' : '$false'}`,
            'if ($attachToRunning) {',
            '  try { $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application") } catch { }',
            '}',
            'if (-not $excel) {',
            '  $excel = New-Object -ComObject Excel.Application',
            '}',
            '$excel.Visible = $true',
            'foreach ($wb in @($excel.Workbooks)) {',
            '  if (($wb.FullName -ieq $targetPath) -or ($wb.Name -ieq $targetName)) { $workbook = $wb; break }',
            '}',
            'if (-not $workbook) {',
            `  $workbook = $excel.Workbooks.Open($targetPath, 0, ${roFlag})`,
            '}',
            '$workbook.Activate()',
            "try { Add-Type -MemberDefinition '[DllImport(\"user32.dll\")] public static extern bool SetForegroundWindow(IntPtr hWnd); [DllImport(\"user32.dll\")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);' -Name XlideWin32 -Namespace XlideHelper } catch { }",
            '[XlideHelper.XlideWin32]::ShowWindow([IntPtr]$excel.Hwnd, 9)',
            '[XlideHelper.XlideWin32]::SetForegroundWindow([IntPtr]$excel.Hwnd)',
        ].join('; ');

        log(`[openWorkbook] Running: powershell -Command "${script}"`);
        const child = cp.spawn('powershell.exe', [
            '-NoProfile',
            '-ExecutionPolicy',
            'Bypass',
            '-Command',
            script,
        ]);
        child.on('spawn', () => {
            log(`[openWorkbook] Spawned powershell.exe (pid=${child.pid ?? 'unknown'})`);
        });
        child.on('error', (err) => {
            log(`[openWorkbook] Error: ${err.message}`);
            void vscode.window.showErrorMessage(`XLIDE: Open Workbook failed: ${err.message}`);
        });
        child.stdout?.on('data', (d: Buffer) => {
            const text = d.toString().trim();
            if (text) {
                log(`[openWorkbook stdout] ${text}`);
            }
        });
        child.stderr?.on('data', (d: Buffer) => {
            const text = d.toString().trim();
            if (text) {
                log(`[openWorkbook stderr] ${text}`);
            }
        });
        child.on('exit', (code, signal) => {
            log(`[openWorkbook] powershell exited with code=${code} signal=${signal ?? 'none'}`);
        });
    }

    function runWindowsExcelMacroReadOnly(filePath: string, macroName: string, attachToRunning: boolean): Promise<void> {
        const script = [
            '$ErrorActionPreference = "Stop"',
            'try {',
            `  $targetPath = ${psSingleQuoted(filePath)}`,
            `  $targetName = ${psSingleQuoted(path.basename(filePath))}`,
            `  $macroName = ${psSingleQuoted(macroName)}`,
            '  $excel = $null',
            '  $workbook = $null',
            `  $attachToRunning = ${attachToRunning ? '$true' : '$false'}`,
            '  if ($attachToRunning) {',
            '    try { $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application") } catch { }',
            '  }',
            '  if (-not $excel) {',
            '    $excel = New-Object -ComObject Excel.Application',
            '  }',
            '  $excel.Visible = $true',
            '  foreach ($wb in @($excel.Workbooks)) {',
            '    if (($wb.FullName -ieq $targetPath) -or ($wb.Name -ieq $targetName)) { $workbook = $wb; break }',
            '  }',
            '  if ($workbook) {',
            '    if (-not $workbook.ReadOnly) {',
            '      throw "REOPEN_BLOCKED|Workbook is already open for editing in Excel. Close it in Excel, then press F5 again so XLIDE can reopen the saved workbook before running the macro."',
            '    }',
            '    try {',
            '      $workbook.Close($false)',
            '      $workbook = $null',
            '    } catch {',
            '      throw ("REOPEN_FAILED|XLIDE could not close the existing read-only workbook before running the macro: " + $_.Exception.Message)',
            '    }',
            '  }',
            '  try {',
            '    $workbook = $excel.Workbooks.Open($targetPath, 0, $true)',
            '  } catch {',
            '    throw ("REOPEN_FAILED|XLIDE could not reopen the workbook. If it is open outside XLIDE, close it in Excel and try again: " + $_.Exception.Message)',
            '  }',
            '  $workbook.Activate()',
            '  try { Add-Type -MemberDefinition \'[DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd); [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);\' -Name XlideWin32 -Namespace XlideHelper } catch { }',
            '  [XlideHelper.XlideWin32]::ShowWindow([IntPtr]$excel.Hwnd, 9)',
            '  [XlideHelper.XlideWin32]::SetForegroundWindow([IntPtr]$excel.Hwnd)',
            '  $macroRef = "\'" + $workbook.Name + "\'!" + $macroName',
            '  try {',
            '    $excel.Run($macroRef)',
            '  } catch {',
            '    throw ("RUN_FAILED|XLIDE could not run the macro: " + $_.Exception.Message)',
            '  }',
            '  [Console]::Out.WriteLine("XLIDE_MACRO_OK")',
            '} catch {',
            '  [Console]::Error.WriteLine("XLIDE_MACRO_ERROR|" + $_.Exception.Message)',
            '  exit 1',
            '}',
        ].join('; ');

        log(`[runMacro] Running: ${macroName}`);
        log(`[runMacro] Script: ${script}`);
        return new Promise<void>((resolve, reject) => {
            const child = cp.spawn('powershell.exe', [
                '-NoProfile',
                '-ExecutionPolicy',
                'Bypass',
                '-Command',
                script,
            ]);
            const stderrLines: string[] = [];
            child.on('spawn', () => {
                log(`[runMacro] Spawned powershell.exe (pid=${child.pid ?? 'unknown'})`);
            });
            child.on('error', (err) => {
                log(`[runMacro] Error: ${err.message}`);
                reject(new Error(`RUN_FAILED|${err.message}`));
            });
            child.stdout?.on('data', (d: Buffer) => {
                const text = d.toString().trim();
                if (text) {
                    log(`[runMacro stdout] ${text}`);
                }
            });
            child.stderr?.on('data', (d: Buffer) => {
                for (const line of d.toString().split('\n')) {
                    const trimmed = line.trimEnd();
                    if (trimmed) {
                        stderrLines.push(trimmed);
                        log(`[runMacro stderr] ${trimmed}`);
                    }
                }
            });
            child.on('exit', (code, signal) => {
                log(`[runMacro] powershell exited with code=${code} signal=${signal ?? 'none'}`);
                if (code === 0) {
                    resolve();
                    return;
                }
                const sentinel = stderrLines.find((line) => line.includes('XLIDE_MACRO_ERROR|'));
                const message = sentinel
                    ? sentinel.slice(sentinel.indexOf('XLIDE_MACRO_ERROR|') + 'XLIDE_MACRO_ERROR|'.length)
                    : stderrLines.join('\n') || `PowerShell exited with code ${code}`;
                reject(new Error(message));
            });
        });
    }

    async function runOwnedReadOnlyExcelTestHost(
        filePath: string,
        tests: readonly VbaTestCase[],
        options: VbaTestRunOptions,
    ): Promise<OwnedReadOnlyExcelHostRunResult> {
        const hostScriptDir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-vba-test-host-'));
        const tempWorkbookPath = path.join(hostScriptDir, path.basename(filePath));
        const hostScriptPath = path.join(hostScriptDir, 'run-vba-tests.ps1');
        const runnerModuleName = `XlideRun${Date.now().toString(36).slice(-8)}`;
        try {
            fs.copyFileSync(filePath, tempWorkbookPath);
            await bridge.call<{ ok?: boolean; signatureDropped?: boolean }>('writeModule', {
                path: tempWorkbookPath,
                module: XLIDE_ASSERT_MODULE_NAME,
                source: XLIDE_ASSERT_MODULE_SOURCE,
                kind: 'standard',
            });
            await bridge.call<{ ok?: boolean; signatureDropped?: boolean }>('writeModule', {
                path: tempWorkbookPath,
                module: runnerModuleName,
                source: buildVbaTestDirectRunnerModule(tests, runnerModuleName),
                kind: 'standard',
            });
            const script = buildOwnedReadOnlyExcelTestHostScript(
                tempWorkbookPath,
                vbaTestHostPlanItems(tests),
                { failFast: options.failFast, runnerModuleName },
            );
            fs.writeFileSync(hostScriptPath, script, 'utf8');
        } catch (err) {
            try {
                fs.rmSync(hostScriptDir, { recursive: true, force: true });
            } catch {
                // Best-effort cleanup only.
            }
            throw err;
        }
        log(`[runVbaTests] Running owned read-only Excel host for ${tests.length} test(s).`);
        log(`[runVbaTests] Temporary workbook path: ${tempWorkbookPath}`);
        log(`[runVbaTests] Host script path: ${hostScriptPath}`);

        return new Promise<OwnedReadOnlyExcelHostRunResult>((resolve) => {
            const child = cp.spawn('powershell.exe', [
                '-NoProfile',
                '-ExecutionPolicy',
                'Bypass',
                '-File',
                hostScriptPath,
            ]);
            const events: VbaTestHostOracleEvent[] = [];
            const stderrLines: string[] = [];
            let stdoutBuffer = '';
            let currentMacro: { excelId: string; qualifiedName: string; timeoutMs: number; startedMs: number } | undefined;
            let currentModalBlocker: Extract<VbaTestHostOracleEvent, { kind: 'modal-blocked' }> | undefined;
            let currentTimer: ReturnType<typeof setTimeout> | undefined;
            let startupTimer: ReturnType<typeof setTimeout> | undefined;
            let cleanupTimer: ReturnType<typeof setTimeout> | undefined;
            let ownedExcelPid: number | undefined;
            let ownedExcelKilled = false;
            let sawWorkbookClosed = false;
            let sawExcelQuit = false;
            let timedOutAfter: string | undefined;
            let settled = false;
            let hostScriptCleaned = false;

            const clearCurrentTimer = () => {
                if (currentTimer) {
                    clearTimeout(currentTimer);
                    currentTimer = undefined;
                }
            };
            const clearStartupTimer = () => {
                if (startupTimer) {
                    clearTimeout(startupTimer);
                    startupTimer = undefined;
                }
            };
            const clearCleanupTimer = () => {
                if (cleanupTimer) {
                    clearTimeout(cleanupTimer);
                    cleanupTimer = undefined;
                }
            };
            const cleanupHostScript = () => {
                if (hostScriptCleaned) {
                    return;
                }
                hostScriptCleaned = true;
                try {
                    fs.rmSync(hostScriptDir, { recursive: true, force: true });
                } catch (err) {
                    const message = err instanceof Error ? err.message : String(err);
                    log(`[runVbaTests] Could not delete temporary host script: ${message}`);
                    setTimeout(() => {
                        try {
                            fs.rmSync(hostScriptDir, { recursive: true, force: true });
                        } catch {
                            // Best-effort cleanup only.
                        }
                    }, 1000);
                }
            };

            const eventResult = (event: VbaTestHostOracleEvent): OwnedReadOnlyExcelHostTestResult | undefined => {
                if (event.kind !== 'macro-finished') {
                    return undefined;
                }
                if (event.outcome === 'passed' || event.outcome === 'failed') {
                    return {
                        outcome: event.outcome,
                        durationMs: event.durationMs ?? 0,
                        message: event.message,
                    };
                }
                if (event.outcome === 'timeout' || event.outcome === 'hung') {
                    return {
                        outcome: 'timeout',
                        durationMs: event.durationMs ?? currentMacro?.timeoutMs ?? DEFAULT_VBA_TEST_TIMEOUT_MS,
                        message: event.message,
                    };
                }
                if (event.outcome === 'modal-blocked') {
                    return {
                        outcome: 'modal-blocked',
                        durationMs: event.durationMs ?? currentMacro?.timeoutMs ?? DEFAULT_VBA_TEST_TIMEOUT_MS,
                        message: event.message,
                    };
                }
                return {
                    outcome: 'runner-error',
                    durationMs: event.durationMs ?? 0,
                    message: event.message,
                };
            };

            const finish = (hostError?: string) => {
                if (settled) {
                    return;
                }
                settled = true;
                clearCurrentTimer();
                clearStartupTimer();
                clearCleanupTimer();
                cleanupHostScript();
                const resultsByName = new Map<string, OwnedReadOnlyExcelHostTestResult>();
                for (const event of events) {
                    const result = eventResult(event);
                    if (result && event.kind === 'macro-finished') {
                        resultsByName.set(event.qualifiedName, result);
                    }
                }
                const oracleIssues = validateVbaTestHostOracleTrace(events);
                for (const issue of oracleIssues) {
                    log(`[runVbaTests oracle] ${issue.code}: ${issue.message}`);
                }
                resolve({
                    events,
                    resultsByName,
                    hostError,
                    timedOutAfter,
                });
            };

            const killOwnedExcel = (reason: OwnedExcelKillReason) => {
                if (!ownedExcelPid || ownedExcelKilled) {
                    return;
                }
                ownedExcelKilled = true;
                log(`[runVbaTests] Killing owned Excel process ${ownedExcelPid} after ${reason}.`);
                cp.spawn('taskkill.exe', ['/PID', String(ownedExcelPid), '/T', '/F']);
                const excelId = currentMacro?.excelId ?? events.find((event) => event.kind === 'excel-created')?.excelId;
                if (excelId) {
                    events.push({ kind: 'excel-killed', excelId, reason });
                }
            };

            const armCleanupWatchdog = (stage: 'workbook-closed' | 'excel-quit') => {
                clearCleanupTimer();
                cleanupTimer = setTimeout(() => {
                    if (settled) {
                        return;
                    }
                    const elapsedDescription = `${DEFAULT_VBA_TEST_CLEANUP_GRACE_MS} ms`;
                    if (!sawExcelQuit) {
                        log(`[runVbaTests] Cleanup watchdog elapsed ${elapsedDescription} after workbook close; killing owned Excel.`);
                        killOwnedExcel('cleanup-failed');
                    } else {
                        log(`[runVbaTests] Cleanup watchdog elapsed ${elapsedDescription} after Excel quit; stopping host script.`);
                    }
                    child.kill();
                    finish();
                }, DEFAULT_VBA_TEST_CLEANUP_GRACE_MS);
                log(`[runVbaTests] Cleanup watchdog armed after ${stage} (${DEFAULT_VBA_TEST_CLEANUP_GRACE_MS} ms).`);
            };

            startupTimer = setTimeout(() => {
                if (settled || currentMacro) {
                    return;
                }
                timedOutAfter = 'test host startup';
                const excelId = events.find((event) => event.kind === 'excel-created')?.excelId ?? 'unknown';
                const message = `Excel test host timed out before starting tests after ${DEFAULT_VBA_TEST_TIMEOUT_MS} ms.`;
                events.push({
                    kind: 'macro-finished',
                    excelId,
                    qualifiedName: 'XLIDE.TestHostStartup',
                    outcome: 'timeout',
                    durationMs: DEFAULT_VBA_TEST_TIMEOUT_MS,
                    message,
                });
                killOwnedExcel('timeout');
                child.kill();
                finish(message);
            }, DEFAULT_VBA_TEST_TIMEOUT_MS);

            const armMacroTimeout = (event: Extract<VbaTestHostOracleEvent, { kind: 'macro-started' }>) => {
                clearStartupTimer();
                clearCurrentTimer();
                const timeoutMs = event.timeoutMs ?? DEFAULT_VBA_TEST_TIMEOUT_MS;
                currentMacro = {
                    excelId: event.excelId,
                    qualifiedName: event.qualifiedName,
                    timeoutMs,
                    startedMs: Date.now(),
                };
                currentTimer = setTimeout(() => {
                    if (!currentMacro || timedOutAfter) {
                        return;
                    }
                    timedOutAfter = currentMacro.qualifiedName;
                    const durationMs = Date.now() - currentMacro.startedMs;
                    const modalBlocker = currentModalBlocker;
                    const modalDetail = modalBlocker
                        ? [modalBlocker.title, modalBlocker.message].filter(Boolean).join(': ')
                        : '';
                    const outcome = modalBlocker ? 'modal-blocked' : 'timeout';
                    const message = modalBlocker
                        ? `Blocked by Excel modal dialog${modalDetail ? ` (${modalDetail})` : ''}.`
                        : `Timed out after ${currentMacro.timeoutMs} ms.`;
                    events.push({
                        kind: 'macro-finished',
                        excelId: currentMacro.excelId,
                        qualifiedName: currentMacro.qualifiedName,
                        outcome,
                        durationMs,
                        message,
                    });
                    killOwnedExcel(modalBlocker ? 'modal-blocked' : 'timeout');
                    child.kill();
                    finish(message);
                }, timeoutMs);
            };

            const modalBlockedMessage = (event: Extract<VbaTestHostOracleEvent, { kind: 'modal-blocked' }>): string => {
                const modalDetail = [event.title, event.message].filter(Boolean).join(': ');
                return `Blocked by Excel modal dialog${modalDetail ? ` (${modalDetail})` : ''}.`;
            };

            const handleEvent = (event: VbaTestHostOracleEvent) => {
                events.push(event);
                if (event.kind === 'excel-created') {
                    ownedExcelPid = event.pid;
                    log(`[runVbaTests host] excel-created pid=${ownedExcelPid ?? 'unknown'}`);
                } else if (event.kind === 'host-phase') {
                    log(`[runVbaTests host] phase ${event.phase} ${event.outcome} (${event.durationMs} ms)`);
                } else if (event.kind === 'macro-started') {
                    log(`[runVbaTests host] macro-started ${event.qualifiedName} timeoutMs=${event.timeoutMs ?? DEFAULT_VBA_TEST_TIMEOUT_MS}`);
                    currentModalBlocker = undefined;
                    armMacroTimeout(event);
                } else if (event.kind === 'modal-detected') {
                    log(`[runVbaTests host] modal-detected ${event.qualifiedName} classification=${event.classification ?? 'unknown'} safeToDismiss=${event.safeToDismiss ?? false}`);
                } else if (event.kind === 'modal-dismissed') {
                    log(`[runVbaTests host] modal-dismissed ${event.qualifiedName} button=${event.button ?? 'unknown'} dismissed=${event.dismissed}`);
                } else if (event.kind === 'modal-blocked') {
                    if (currentMacro &&
                        currentMacro.excelId === event.excelId &&
                        currentMacro.qualifiedName === event.qualifiedName) {
                        const macro = currentMacro;
                        currentModalBlocker = event;
                        const durationMs = Date.now() - macro.startedMs;
                        const message = modalBlockedMessage(event);
                        events.push({
                            kind: 'macro-finished',
                            excelId: macro.excelId,
                            qualifiedName: macro.qualifiedName,
                            outcome: 'modal-blocked',
                            durationMs,
                            message,
                        });
                        killOwnedExcel('modal-blocked');
                        child.kill();
                        finish(message);
                    }
                    log(`[runVbaTests host] modal-blocked ${event.qualifiedName}: ${event.reason}`);
                } else if (event.kind === 'macro-finished') {
                    log(`[runVbaTests host] macro-finished ${event.qualifiedName} outcome=${event.outcome}`);
                    clearCurrentTimer();
                    currentMacro = undefined;
                    currentModalBlocker = undefined;
                } else if (event.kind === 'workbook-closed') {
                    sawWorkbookClosed = true;
                    log(`[runVbaTests host] workbook-closed durationMs=${event.durationMs ?? 'unknown'}`);
                    armCleanupWatchdog('workbook-closed');
                } else if (event.kind === 'excel-quit') {
                    sawExcelQuit = true;
                    log(`[runVbaTests host] excel-quit durationMs=${event.durationMs ?? 'unknown'}`);
                    if (sawWorkbookClosed) {
                        armCleanupWatchdog('excel-quit');
                    }
                }
            };

            const handleStdoutText = (text: string) => {
                stdoutBuffer += text;
                const lines = stdoutBuffer.split(/\r?\n/);
                stdoutBuffer = lines.pop() ?? '';
                for (const line of lines) {
                    const trimmed = line.trim();
                    if (!trimmed) {
                        continue;
                    }
                    try {
                        const event = parseVbaTestHostEventLine(trimmed);
                        if (event) {
                            handleEvent(event);
                        } else {
                            log(`[runVbaTests host stdout] ${trimmed}`);
                        }
                    } catch (err) {
                        log(`[runVbaTests host stdout] ${trimmed}`);
                        log(`[runVbaTests host parse] ${err instanceof Error ? err.message : String(err)}`);
                    }
                }
            };

            child.on('spawn', () => {
                log(`[runVbaTests] Spawned owned host powershell.exe (pid=${child.pid ?? 'unknown'})`);
            });
            child.on('error', (err) => {
                const message = `RUNNER_FAILED|${err.message}`;
                if (currentMacro) {
                    events.push({
                        kind: 'macro-finished',
                        excelId: currentMacro.excelId,
                        qualifiedName: currentMacro.qualifiedName,
                        outcome: 'runner-error',
                        durationMs: Date.now() - currentMacro.startedMs,
                        message,
                    });
                }
                killOwnedExcel('runner-error');
                finish(message);
            });
            child.stdout?.on('data', (data: Buffer) => {
                handleStdoutText(data.toString());
            });
            child.stderr?.on('data', (data: Buffer) => {
                for (const line of data.toString().split(/\r?\n/)) {
                    const trimmed = line.trimEnd();
                    if (trimmed) {
                        stderrLines.push(trimmed);
                        log(`[runVbaTests host stderr] ${trimmed}`);
                    }
                }
            });
            child.on('exit', (code, signal) => {
                if (stdoutBuffer.trim()) {
                    handleStdoutText('\n');
                }
                log(`[runVbaTests] owned host powershell exited with code=${code} signal=${signal ?? 'none'}`);
                if (settled) {
                    return;
                }
                if (code === 0) {
                    finish();
                    return;
                }
                const sentinel = stderrLines.find((line) => line.includes('XLIDE_TEST_HOST_ERROR|'));
                const hostError = sentinel
                    ? sentinel.slice(sentinel.indexOf('XLIDE_TEST_HOST_ERROR|') + 'XLIDE_TEST_HOST_ERROR|'.length)
                    : stderrLines.join('\n') || `PowerShell exited with code ${code}`;
                if (currentMacro) {
                    events.push({
                        kind: 'macro-finished',
                        excelId: currentMacro.excelId,
                        qualifiedName: currentMacro.qualifiedName,
                        outcome: 'runner-error',
                        durationMs: Date.now() - currentMacro.startedMs,
                        message: hostError,
                    });
                }
                killOwnedExcel('runner-error');
                finish(hostError);
            });
        });
    }

    async function runWorkbookVbaTests(
        filePath: string,
        progress?: vscode.Progress<{ message?: string; increment?: number }>,
        options: VbaTestRunOptions = {},
    ): Promise<VbaTestRunExecution> {
        const startedAt = new Date();
        const startedMs = Date.now();
        progress?.report({ message: 'Discovering tests...' });
        const discovery = await discoverWorkbookVbaTests(bridge, filePath, options.selection);
        const results: VbaTestRunItem[] = [];

        if (discovery.tests.length === 0) {
            return {
                report: createVbaTestRunReport({
                    filePath,
                    startedAt,
                    durationMs: Date.now() - startedMs,
                    discovery,
                    results,
                }),
                hostEvents: [],
            };
        }

        if (process.platform !== 'win32') {
            const message = 'VBA test execution currently requires Excel COM on Windows.';
            for (const test of discovery.tests) {
                results.push({
                    test,
                    status: 'skipped',
                    durationMs: 0,
                    error: message,
                });
            }
            return {
                report: createVbaTestRunReport({
                    filePath,
                    startedAt,
                    durationMs: Date.now() - startedMs,
                    discovery,
                    results,
                }),
                hostEvents: [],
            };
        }

        log('[runVbaTests] attachToRunningExcel=false (owned read-only test host)');
        log(`[runVbaTests] selection=${describeVbaTestSelection(options.selection) || 'all tests'} failFast=${Boolean(options.failFast)}`);
        const executableTests = discovery.tests.filter((test) => !test.metadata.skipReason);
        progress?.report({ message: `Running ${executableTests.length} test(s) in owned Excel...` });
        const hostRun: OwnedReadOnlyExcelHostRunResult = executableTests.length > 0
            ? await runOwnedReadOnlyExcelTestHost(filePath, executableTests, options)
            : {
                events: [],
                resultsByName: new Map<string, OwnedReadOnlyExcelHostTestResult>(),
            };
        let stoppedAfter: string | undefined;

        for (const test of discovery.tests) {
            if (stoppedAfter) {
                const message = `Not run because fail-fast stopped after ${stoppedAfter}.`;
                results.push({ test, status: 'skipped', durationMs: 0, error: message });
                log(`[runVbaTests] SKIP ${test.qualifiedName}: ${message}`);
                continue;
            }

            if (test.metadata.skipReason) {
                results.push({
                    test,
                    status: 'skipped',
                    durationMs: 0,
                    error: test.metadata.skipReason,
                });
                log(`[runVbaTests] SKIP ${test.qualifiedName}: ${test.metadata.skipReason}`);
                continue;
            }

            const hostResult = hostRun.resultsByName.get(test.qualifiedName);
            let result: VbaTestRunItem;
            if (hostResult) {
                result = vbaTestRunItemFromHostResult(test, hostResult);
            } else if (hostRun.hostError) {
                result = { test, status: 'host-error', durationMs: 0, error: hostRun.hostError };
            } else if (hostRun.timedOutAfter) {
                result = {
                    test,
                    status: 'skipped',
                    durationMs: 0,
                    error: `Not run because test host timed out after ${hostRun.timedOutAfter}.`,
                };
            } else {
                result = {
                    test,
                    status: 'host-error',
                    durationMs: 0,
                    error: 'The Excel test host did not emit a result for this test.',
                };
            }
            results.push(result);
            logVbaTestRunItem(result);
            if (shouldFailFastStop(result, options.failFast)) {
                stoppedAfter = test.qualifiedName;
            }
        }

        return {
            report: createVbaTestRunReport({
                filePath,
                startedAt,
                durationMs: Date.now() - startedMs,
                discovery,
                results,
            }),
            hostEvents: hostRun.events,
        };
    }

    function vbaTestRunItemFromHostResult(
        test: VbaTestCase,
        hostResult: OwnedReadOnlyExcelHostTestResult,
    ): VbaTestRunItem {
        const message = hostResult.message ? vbaTestFailureMessage(new Error(hostResult.message)) : undefined;
        if (hostResult.outcome === 'passed') {
            if (test.metadata.xfailReason) {
                return {
                    test,
                    status: 'xpass',
                    durationMs: hostResult.durationMs,
                    error: `Expected failure did not occur: ${test.metadata.xfailReason}`,
                };
            }
            return { test, status: 'passed', durationMs: hostResult.durationMs };
        }
        if (hostResult.outcome === 'failed') {
            if (test.metadata.xfailReason) {
                return { test, status: 'xfail', durationMs: hostResult.durationMs, error: message };
            }
            return { test, status: 'failed', durationMs: hostResult.durationMs, error: message };
        }
        if (hostResult.outcome === 'timeout') {
            return {
                test,
                status: 'timeout',
                durationMs: hostResult.durationMs,
                error: message ?? `Timed out after ${hostResult.durationMs} ms.`,
            };
        }
        if (hostResult.outcome === 'modal-blocked') {
            return {
                test,
                status: 'host-error',
                durationMs: hostResult.durationMs,
                error: message ?? 'Blocked by an Excel modal dialog.',
            };
        }
        return {
            test,
            status: 'host-error',
            durationMs: hostResult.durationMs,
            error: message ?? 'The Excel test host failed while running this test.',
        };
    }

    function logVbaTestRunItem(result: VbaTestRunItem): void {
        const detail = result.error ? `: ${result.error}` : '';
        log(`[runVbaTests] ${result.status.toUpperCase()} ${result.test.qualifiedName} (${result.durationMs} ms)${detail}`);
    }

    function shouldFailFastStop(result: VbaTestRunItem, failFast?: boolean): boolean {
        return Boolean(failFast) &&
            (result.status === 'failed' ||
                result.status === 'xpass' ||
                result.status === 'timeout' ||
                result.status === 'host-error');
    }

    function workbookStateKey(filePath: string): string {
        const normalized = path.normalize(filePath);
        return process.platform === 'win32' ? normalized.toLowerCase() : normalized;
    }

    function updateLastFailedVbaTestRun(report: VbaTestRunReport): void {
        const failed = report.results
            .filter((result) => VBA_TEST_RERUN_FAILED_STATUSES.has(result.status))
            .map((result) => ({
                id: result.test.id,
                qualifiedName: result.test.qualifiedName,
                status: result.status,
            }));
        const key = workbookStateKey(report.filePath);
        if (failed.length === 0) {
            lastFailedVbaTestRuns.delete(key);
            return;
        }
        lastFailedVbaTestRuns.set(key, {
            testIds: failed.map((test) => test.id),
            tests: failed,
        });
    }

    function lastFailedRunForWorkbook(filePath: string): VbaTestLastFailedRun | undefined {
        return lastFailedVbaTestRuns.get(workbookStateKey(filePath));
    }

    async function rerunFailedVbaTestsForWorkbook(filePath: string): Promise<void> {
        const failed = lastFailedRunForWorkbook(filePath);
        if (!failed || failed.testIds.length === 0) {
            void vscode.window.showInformationMessage(
                `XLIDE: No failed VBA tests to rerun for "${path.basename(filePath)}".`,
            );
            return;
        }
        await runVbaTestsForWorkbook(filePath, {
            selection: {
                testIds: failed.testIds,
            },
        });
    }

    async function runVbaTestsForWorkbook(
        filePath: string,
        options: VbaTestRunOptions = {},
    ): Promise<void> {
        const name = path.basename(filePath);
        const selectionDescription = describeVbaTestSelection(options.selection);
        const runScope = selectionDescription ? ` (${selectionDescription})` : '';
        let resultsPanelRunning = false;
        try {
            const support = await vbaTestSupportStatus(filePath);
            if (!support.canRun) {
                openVbaTestsForWorkbook(filePath);
                void vscode.window.showWarningMessage(
                    `XLIDE: Install XlideAssert.bas from the Unit Tests GUI before running tests for "${name}".`,
                );
                return;
            }
            const runtime = await checkExcelComAvailability();
            if (!runtime.canRun) {
                openVbaTestsForWorkbook(filePath);
                void vscode.window.showWarningMessage(`XLIDE: ${runtime.description}`);
                return;
            }
            resultsPanelRunning = true;
            setVbaTestResultsRunning(filePath, true);
            let execution: VbaTestRunExecution | undefined;
            await vscode.window.withProgress(
                {
                    location: vscode.ProgressLocation.Notification,
                    title: `XLIDE: Running VBA tests${runScope} for "${name}"...`,
                    cancellable: false,
                },
                async (progress) => {
                    execution = await runWorkbookVbaTests(filePath, progress, options);
                },
            );
            if (!execution) {
                return;
            }
            const { report, hostEvents } = execution;
            log(`[runVbaTests] Report JSON:\n${JSON.stringify(report, null, 2)}`);
            try {
                const testSettings = await effectiveWorkbookTestSettings(filePath);
                const artifacts = await writeVbaTestRunArtifacts(report, hostEvents, {
                    outputFolder: testSettings.artifactFolder,
                    retention: testSettings.artifactRetention,
                });
                log(`[runVbaTests] Artifacts written to ${artifacts.runDirectory}`);
                log(`[runVbaTests] CI status written to ${artifacts.statusPath}`);
                log(`[runVbaTests] Artifact settings source folder=${testSettings.artifactFolderSource} retention=${testSettings.artifactRetentionSource}`);
            } catch (err) {
                const message = err instanceof Error ? err.message : String(err);
                log(`[runVbaTests] Artifact write failed: ${message}`);
                void vscode.window.showWarningMessage(`XLIDE: VBA test artifacts could not be written: ${message}`);
            }
            updateLastFailedVbaTestRun(report);
            openVbaTestResults(context, report, {
                onRerunFailed: () => rerunFailedVbaTestsForWorkbook(filePath),
            });
            showVbaTestRunOutcome(report);
        } catch (err) {
            const msg = err instanceof Error ? err.message : String(err);
            log(`[runVbaTests] FAILED: ${msg}`);
            vscode.window.showErrorMessage(`XLIDE: VBA tests failed: ${msg}`);
        } finally {
            if (resultsPanelRunning) {
                setVbaTestResultsRunning(filePath, false);
            }
        }
    }

    function openVbaTestsForWorkbook(filePath: string): void {
        openVbaTestsPanel(context, filePath, {
            getModel: () => vbaTestsPanelModel(filePath),
            onInstallSupport: async () => {
                await installVbaTestSupportModule(filePath);
            },
            onRunAll: async () => {
                await runVbaTestsForWorkbook(filePath);
            },
            onRunWithFilters: async (filters) => {
                await runVbaTestsForWorkbook(filePath, vbaTestRunOptionsFromFilters(filters));
            },
            onRerunFailed: async () => {
                await rerunFailedVbaTestsForWorkbook(filePath);
            },
            onDidChangeWorkbookTree: explorer.onDidChangeTreeData,
        });
    }

    async function vbaTestsPanelModel(filePath: string): Promise<VbaTestsPanelModel> {
        const [support, runtime, discovery] = await Promise.all([
            vbaTestSupportStatus(filePath),
            checkExcelComAvailability(),
            vbaTestsDiscoveryStatus(filePath),
        ]);
        const lastFailed = lastFailedRunForWorkbook(filePath);
        return {
            filePath,
            workbookName: path.basename(filePath),
            support,
            runtime,
            discovery,
            lastFailed: lastFailed
                ? {
                    count: lastFailed.testIds.length,
                    tests: lastFailed.tests,
                }
                : undefined,
        };
    }

    async function vbaTestsDiscoveryStatus(filePath: string): Promise<VbaTestsPanelModel['discovery']> {
        try {
            const discovery = await discoverWorkbookVbaTests(bridge, filePath);
            const tags = summarizeVbaTestTags(discovery.tests);
            const taggedTests = discovery.tests.filter((test) => test.metadata.tags.length > 0).length;
            return {
                totalTests: discovery.tests.length,
                taggedTests,
                untaggedTests: Math.max(0, discovery.tests.length - taggedTests),
                tags,
            };
        } catch (err) {
            const error = err instanceof Error ? err.message : String(err);
            return {
                totalTests: 0,
                taggedTests: 0,
                untaggedTests: 0,
                tags: [],
                error,
            };
        }
    }

    function vbaTestRunOptionsFromFilters(filters: VbaTestsRunFilterRequest): VbaTestRunOptions {
        const selection: VbaTestSelectionOptions = {};
        if (filters.includeTags.length > 0) {
            selection.includeTags = filters.includeTags;
        }
        if (filters.excludeTags.length > 0) {
            selection.excludeTags = filters.excludeTags;
        }
        return {
            selection,
            failFast: filters.failFast,
        };
    }

    async function vbaTestSupportStatus(filePath: string): Promise<VbaTestSupportStatusModel> {
        try {
            const modules = await bridge.call<Array<{ name: string; type: string }>>(
                'listModules',
                { path: filePath },
            );
            const existing = modules.find(
                (module) => module.name.toLowerCase() === XLIDE_ASSERT_MODULE_NAME.toLowerCase(),
            );
            if (!existing) {
                return {
                    state: 'missing',
                    title: 'XlideAssert.bas Not Installed',
                    description: 'The bundled test support module must be installed before XLIDE can run workbook tests.',
                    actionLabel: 'Install',
                    canInstall: true,
                    canRun: false,
                };
            }
            if (existing.type !== 'standard') {
                return {
                    state: 'blocked',
                    title: `${XLIDE_ASSERT_MODULE_NAME} Name Conflict`,
                    description: `"${XLIDE_ASSERT_MODULE_NAME}" exists as a ${existing.type} module. Rename it before installing the XLIDE test support module.`,
                    actionLabel: 'Blocked',
                    canInstall: false,
                    canRun: false,
                };
            }

            const current = await bridge.call<{ source?: string } | string>(
                'readModule',
                { path: filePath, module: existing.name },
            );
            const installed = normalizeVbaTestSupportModuleSource(moduleSourceFromReadResult(current)) ===
                normalizeVbaTestSupportModuleSource(XLIDE_ASSERT_MODULE_SOURCE);
            if (installed) {
                return {
                    state: 'installed',
                    title: 'XlideAssert.bas Installed',
                    description: 'Workbook tests can run through the XLIDE-owned read-only Excel test host.',
                    actionLabel: 'Installed',
                    canInstall: false,
                    canRun: true,
                };
            }
            return {
                state: 'outdated',
                title: 'XlideAssert.bas Needs Update',
                description: 'The workbook has an XlideAssert standard module, but it does not match the bundled XLIDE test support module.',
                actionLabel: 'Update',
                canInstall: true,
                canRun: false,
            };
        } catch (err) {
            const message = err instanceof Error ? err.message : String(err);
            return {
                state: 'unknown',
                title: 'Test Support Unknown',
                description: `XLIDE could not inspect the workbook test support module: ${message}`,
                actionLabel: 'Refresh',
                canInstall: false,
                canRun: false,
            };
        }
    }

    function moduleSourceFromReadResult(result: { source?: string } | string): string {
        return typeof result === 'string' ? result : result.source ?? '';
    }

    function showVbaTestRunOutcome(report: VbaTestRunReport): void {
        const summary = summarizeVbaTestRun(report);
        const label = path.basename(report.filePath);
        if (summary.total === 0) {
            const selectionDescription = describeVbaTestSelection(report.discovery.selection);
            if (selectionDescription && report.discovery.unfilteredTestCount > 0) {
                void vscode.window.showInformationMessage(
                    `XLIDE: No VBA tests matched ${selectionDescription} in "${label}".`,
                );
                return;
            }
            void vscode.window.showInformationMessage(
                `XLIDE: No VBA tests were discovered in "${label}". Add '@xlide-test' above a no-argument standard-module Sub.`,
            );
            return;
        }
        if (summary.failed > 0 || summary.xpass > 0 || summary.timeout > 0 || summary.hostError > 0) {
            void vscode.window.showWarningMessage(
                `XLIDE: "${label}" VBA tests finished with ${summary.failed} failed, ${summary.timeout} timed out, ${summary.hostError} host error(s), ${summary.xpass} unexpected passed, ${summary.passed} passed, ${summary.skipped} skipped.`,
            );
            return;
        }
        if (summary.skipped > 0) {
            void vscode.window.showWarningMessage(
                `XLIDE: "${label}" VBA tests were discovered but ${summary.skipped} were skipped.`,
            );
            return;
        }
        void vscode.window.showInformationMessage(
            `XLIDE: "${label}" passed ${summary.passed} VBA test(s).`,
        );
    }

    function refreshVbaProjectState(filePath: string): void {
        vbaIndex.invalidate(filePath);
        invalidateVbaMemberCompletionCache(filePath);
        explorer.refresh();
    }

    async function installVbaTestSupportModule(filePath: string): Promise<boolean> {
        const modules = await bridge.call<Array<{ name: string; type: string }>>(
            'listModules',
            { path: filePath },
        );
        const existing = modules.find(
            (module) => module.name.toLowerCase() === XLIDE_ASSERT_MODULE_NAME.toLowerCase(),
        );
        if (existing && existing.type !== 'standard') {
            void vscode.window.showErrorMessage(
                `XLIDE: "${XLIDE_ASSERT_MODULE_NAME}" already exists as a ${existing.type} module. Rename it before installing the test support module.`,
            );
            return false;
        }
        if (existing) {
            const current = await bridge.call<{ source: string }>(
                'readModule',
                { path: filePath, module: existing.name },
            );
            if (
                normalizeVbaTestSupportModuleSource(current.source) ===
                normalizeVbaTestSupportModuleSource(XLIDE_ASSERT_MODULE_SOURCE)
            ) {
                void vscode.window.showInformationMessage(
                    `XLIDE: "${XLIDE_ASSERT_MODULE_NAME}" is already installed in "${path.basename(filePath)}".`,
                );
                return false;
            }
            const choice = await vscode.window.showWarningMessage(
                `Update the existing "${XLIDE_ASSERT_MODULE_NAME}" module in "${path.basename(filePath)}"?`,
                { modal: true },
                'Update',
            );
            if (choice !== 'Update') {
                return false;
            }
        }

        const result = await bridge.call<{ ok?: boolean; signatureDropped?: boolean }>(
            'writeModule',
            {
                path: filePath,
                module: XLIDE_ASSERT_MODULE_NAME,
                source: XLIDE_ASSERT_MODULE_SOURCE,
                kind: 'standard',
            },
        );
        notifySignatureDropped(filePath, Boolean(result.signatureDropped));
        const summaryText = logChangeSummary('installVbaTestSupport', {
            operation: existing ? 'Update VBA test support module' : 'Install VBA test support module',
            changed: [XLIDE_ASSERT_MODULE_NAME],
        });
        recordWriteAudit({
            command: 'xlide.installVbaTestSupport',
            operation: 'write-module',
            outcome: 'succeeded',
            workbookPath: filePath,
            moduleName: XLIDE_ASSERT_MODULE_NAME,
            summary: summaryText,
        });
        refreshVbaProjectState(filePath);
        void vscode.window.showInformationMessage(
            `XLIDE: "${XLIDE_ASSERT_MODULE_NAME}" ${existing ? 'updated' : 'installed'} in "${path.basename(filePath)}".`,
        );
        return true;
    }

    function resolveWorkbookPath(node?: XlideNode): string | undefined {
        let filePath = node?.filePath;
        if (!filePath) {
            const active = vscode.window.activeTextEditor;
            if (active && active.document.uri.scheme === XLIDE_SCHEME) {
                filePath = decodeModuleUri(active.document.uri).xlsmPath;
            }
        }
        return filePath;
    }

    function activeLocalVbaEditor(): vscode.TextEditor | undefined {
        const editor = vscode.window.activeTextEditor;
        if (!editor || editor.document.uri.scheme !== XLIDE_SCHEME || editor.document.uri.authority) {
            return undefined;
        }
        return editor;
    }

    function procedureNameAtCursor(editor: vscode.TextEditor): string | undefined {
        const source = editor.document.getText();
        const offset = editor.document.offsetAt(editor.selection.active);
        const parsed = parseModule(source);
        const procedure = parsed.members
            .filter((member): member is ProcedureMember => member.kind === 'Procedure')
            .filter((member) => spanContainsOffset(member.span, offset))
            .sort((left, right) => spanLength(left.span) - spanLength(right.span))[0];
        return procedure?.name;
    }

    function parseTagFilterInput(value: string | undefined): string[] {
        return [...new Set((value ?? '')
            .split(',')
            .map((tag) => tag.trim())
            .filter((tag) => tag.length > 0))];
    }

    async function promptVbaTestRunOptions(): Promise<VbaTestRunOptions | undefined> {
        const includeInput = await vscode.window.showInputBox({
            title: 'XLIDE: Include VBA Test Tags',
            prompt: 'Comma-separated tags to include. Leave blank to include all tags.',
            placeHolder: 'smoke, fast',
        });
        if (includeInput === undefined) {
            return undefined;
        }

        const excludeInput = await vscode.window.showInputBox({
            title: 'XLIDE: Exclude VBA Test Tags',
            prompt: 'Comma-separated tags to exclude. Leave blank to exclude none.',
            placeHolder: 'slow, external',
        });
        if (excludeInput === undefined) {
            return undefined;
        }

        const modeOptions: Array<{ label: string; description: string; failFast: boolean }> = [
            {
                label: 'Run All Matches',
                description: 'Continue after failures and unexpected passes.',
                failFast: false,
            },
            {
                label: 'Fail Fast',
                description: 'Stop after the first failure or unexpected pass.',
                failFast: true,
            },
        ];
        const mode = await vscode.window.showQuickPick(modeOptions, {
            title: 'XLIDE: VBA Test Run Mode',
            placeHolder: 'Choose how XLIDE handles failures.',
        });
        if (!mode) {
            return undefined;
        }

        const includeTags = parseTagFilterInput(includeInput);
        const excludeTags = parseTagFilterInput(excludeInput);
        const selection: VbaTestSelectionOptions = {};
        if (includeTags.length > 0) {
            selection.includeTags = includeTags;
        }
        if (excludeTags.length > 0) {
            selection.excludeTags = excludeTags;
        }
        return {
            selection,
            failFast: mode.failFast,
        };
    }

    async function runCurrentModuleVbaTests(): Promise<void> {
        const editor = activeLocalVbaEditor();
        if (!editor) {
            vscode.window.showWarningMessage('XLIDE: Open a local workbook VBA module to run module tests.');
            return;
        }
        if (editor.document.isDirty) {
            await editor.document.save();
        }
        const { xlsmPath, moduleName } = decodeModuleUri(editor.document.uri);
        await runVbaTestsForWorkbook(xlsmPath, {
            selection: { moduleName },
        });
    }

    async function runVbaTestAtCursor(): Promise<void> {
        const editor = activeLocalVbaEditor();
        if (!editor) {
            vscode.window.showWarningMessage('XLIDE: Open a local workbook VBA module to run the test at the cursor.');
            return;
        }

        const procedureName = procedureNameAtCursor(editor);
        if (!procedureName) {
            vscode.window.showWarningMessage('XLIDE: Cursor is not inside a VBA procedure.');
            return;
        }
        if (editor.document.isDirty) {
            await editor.document.save();
        }
        const { xlsmPath, moduleName } = decodeModuleUri(editor.document.uri);
        await runVbaTestsForWorkbook(xlsmPath, {
            selection: {
                moduleName,
                procedureName,
            },
        });
    }

    async function resolveWorkbookExportFolder(
        filePath: string,
        openLabel: string,
    ): Promise<ResolvedModuleSyncSettings | undefined> {
        const existing = await effectiveWorkbookModuleSyncSettings(filePath);
        if (existing.folderPath) {
            return {
                folderPath: existing.folderPath,
                folderPathSource: existing.folderPathSource,
                exportMode: existing.exportMode,
                exportModeSource: existing.exportModeSource,
                settingsPath: existing.settingsPath,
            };
        }

        const selected = await vscode.window.showOpenDialog({
            canSelectFiles: false,
            canSelectFolders: true,
            canSelectMany: false,
            openLabel,
            defaultUri: vscode.Uri.file(path.dirname(filePath)),
        });
        if (!selected || selected.length === 0) {
            return undefined;
        }
        return {
            folderPath: selected[0].fsPath,
            folderPathSource: 'session',
            exportMode: existing.exportMode,
            exportModeSource: existing.exportModeSource,
            settingsPath: existing.settingsPath,
        };
    }

    async function resolveWorkbookImportFolder(
        filePath: string,
    ): Promise<ResolvedModuleSyncSettings | undefined> {
        const existing = await effectiveWorkbookModuleSyncSettings(filePath);
        if (existing.folderPath) {
            return {
                folderPath: existing.folderPath,
                folderPathSource: existing.folderPathSource,
                importMode: existing.importMode,
                importModeSource: existing.importModeSource,
                settingsPath: existing.settingsPath,
            };
        }

        const selected = await vscode.window.showOpenDialog({
            canSelectFiles: false,
            canSelectFolders: true,
            canSelectMany: false,
            openLabel: 'Select folder to import from',
            defaultUri: vscode.Uri.file(path.dirname(filePath)),
        });
        return selected?.[0]?.fsPath ? {
            folderPath: selected[0].fsPath,
            folderPathSource: 'session',
            importMode: existing.importMode,
            importModeSource: existing.importModeSource,
            settingsPath: existing.settingsPath,
        } : undefined;
    }

    async function resolveWorkbookExportFolderFromSettings(
        filePath: string,
    ): Promise<ResolvedModuleSyncSettings | undefined> {
        const existing = await effectiveWorkbookModuleSyncSettings(filePath);
        if (!existing.folderPath) {
            return undefined;
        }
        return {
            folderPath: existing.folderPath,
            folderPathSource: existing.folderPathSource,
            exportMode: existing.exportMode,
            exportModeSource: existing.exportModeSource,
            settingsPath: existing.settingsPath,
        };
    }

    async function resolveWorkbookImportFolderFromSettings(
        filePath: string,
    ): Promise<ResolvedModuleSyncSettings | undefined> {
        const existing = await effectiveWorkbookModuleSyncSettings(filePath);
        if (!existing.folderPath) {
            return undefined;
        }
        return {
            folderPath: existing.folderPath,
            folderPathSource: existing.folderPathSource,
            importMode: existing.importMode,
            importModeSource: existing.importModeSource,
            settingsPath: existing.settingsPath,
        };
    }

    async function chooseModuleSyncFolder(
        filePath: string,
        currentFolder: string | undefined,
        openLabel: string,
    ): Promise<string | undefined> {
        const selected = await vscode.window.showOpenDialog({
            canSelectFiles: false,
            canSelectFolders: true,
            canSelectMany: false,
            openLabel,
            defaultUri: currentFolder
                ? vscode.Uri.file(currentFolder)
                : vscode.Uri.file(path.dirname(filePath)),
        });
        return selected?.[0]?.fsPath;
    }

    function syncSettingsFromPlan(plan: ModuleSyncPlan): ModuleSyncSettings {
        return {
            folderPath: plan.folderPath,
            folderPathSource: plan.folderPathSource,
            exportMode: plan.exportMode,
            exportModeSource: plan.exportModeSource,
            importMode: plan.importMode,
            importModeSource: plan.importModeSource,
            settingsPath: plan.settingsPath,
        };
    }

    async function buildExportSyncPlanFromSettings(
        filePath: string,
        settings: ModuleSyncSettings,
    ): Promise<ModuleSyncPlan> {
        log(`[exportModules] Target folder: ${settings.folderPath}`);
        log(`[exportModules] Mode: ${settings.exportMode ?? 'exportAll'}`);
        return buildExportModuleSyncPlan(bridge, {
            workbookPath: filePath,
            exportFolder: settings.folderPath,
            exportMode: settings.exportMode,
            folderPathSource: settings.folderPathSource,
            exportModeSource: settings.exportModeSource,
            settingsPath: settings.settingsPath,
        });
    }

    async function buildImportSyncPlanFromSettings(
        filePath: string,
        settings: ModuleSyncSettings,
    ): Promise<ModuleSyncPlan> {
        log(`[importModules] Source folder: ${settings.folderPath}`);
        log(`[importModules] Mode: ${settings.importMode ?? 'updateOnly'}`);
        return buildImportModuleSyncPlan(bridge, {
            workbookPath: filePath,
            importFolder: settings.folderPath,
            importMode: settings.importMode,
            folderPathSource: settings.folderPathSource,
            importModeSource: settings.importModeSource,
            settingsPath: settings.settingsPath,
        });
    }

    async function buildExportSyncPlanFromWorkbookSettings(
        filePath: string,
    ): Promise<ModuleSyncPlan | undefined> {
        const settings = await resolveWorkbookExportFolderFromSettings(filePath);
        return settings ? buildExportSyncPlanFromSettings(filePath, settings) : undefined;
    }

    async function buildImportSyncPlanFromWorkbookSettings(
        filePath: string,
    ): Promise<ModuleSyncPlan | undefined> {
        const settings = await resolveWorkbookImportFolderFromSettings(filePath);
        return settings ? buildImportSyncPlanFromSettings(filePath, settings) : undefined;
    }

    async function persistModuleSyncSettings(
        filePath: string,
        settings: ModuleSyncSettings,
    ): Promise<string> {
        const updated = await updateWorkbookModuleSyncSettings(filePath, {
            folderPath: settings.folderPath,
            exportMode: settings.exportMode,
            importMode: settings.importMode,
        });
        return updated.settingsPath;
    }

    async function saveModuleSyncSettings(
        filePath: string,
        command: string,
        settings: ModuleSyncSettings,
    ): Promise<ModuleSyncApplyResult> {
        const configPath = await persistModuleSyncSettings(filePath, settings);
        const summary = 'Sync settings: 1 changed';
        log(`[moduleSyncSettings] Config updated: ${configPath}`);
        recordWriteAudit({
            command,
            operation: 'configure-module-sync',
            outcome: 'succeeded',
            workbookPath: filePath,
            targetPath: settings.folderPath,
            summary,
        });
        return {
            summary,
            changed: 1,
            skipped: 0,
            failed: 0,
        };
    }

    async function exportActiveModule(): Promise<void> {
        const editor = vscode.window.activeTextEditor;
        if (!editor || editor.document.uri.scheme !== XLIDE_SCHEME || editor.document.uri.authority) {
            vscode.window.showWarningMessage('XLIDE: Open a local workbook VBA module to export the current module.');
            return;
        }

        if (editor.document.isDirty) {
            const saved = await editor.document.save();
            if (!saved) {
                vscode.window.showWarningMessage('XLIDE: Save the current module before exporting it.');
                return;
            }
        }

        const { xlsmPath, moduleName } = decodeModuleUri(editor.document.uri);
        const target = await resolveWorkbookExportFolder(xlsmPath, 'Select export folder');
        if (!target) {
            return;
        }

        log(`[exportCurrentModule] Workbook: ${xlsmPath}`);
        log(`[exportCurrentModule] Module: ${moduleName}`);
        log(`[exportCurrentModule] Target folder: ${target.folderPath}`);
        log(`[exportCurrentModule] Mode: ${target.exportMode}`);

        const result = await exportWorkbookModule(bridge, {
            filePath: xlsmPath,
            moduleName,
            exportFolder: target.folderPath,
            exportMode: target.exportMode,
        });
        const changeSummary: XlideChangeSummary = {
            operation: 'Export current module',
            changed: result.writtenFiles,
        };
        const summaryText = logChangeSummary('exportCurrentModule', changeSummary);
        recordWriteAudit({
            command: 'xlide.exportCurrentModuleToFolder',
            operation: 'export-current-module',
            outcome: 'succeeded',
            workbookPath: xlsmPath,
            moduleName,
            targetPath: target.folderPath,
            summary: summaryText,
        });

        log(`[exportCurrentModule] Config updated: ${result.configPath}`);
        vscode.window.showInformationMessage(
            `XLIDE: ${summaryText} [mode=${result.exportMode}]`,
        );
    }

    async function showExportModulesDiffGui(filePath: string): Promise<void> {
        const target = await resolveWorkbookExportFolder(filePath, 'Select export folder');
        if (!target) {
            return;
        }

        log(`[exportModules] Workbook: ${filePath}`);
        log(`[exportModules] Target folder: ${target.folderPath}`);
        log(`[exportModules] Mode: ${target.exportMode}`);

        const plan = await buildExportSyncPlanFromSettings(filePath, {
            folderPath: target.folderPath,
            folderPathSource: target.folderPathSource,
            exportMode: target.exportMode,
            exportModeSource: target.exportModeSource,
            settingsPath: target.settingsPath,
        });
        const result = await openModuleSyncPreview(
            context,
            plan,
            (currentPlan, selectedIds) => applyExportModuleSyncPlan(currentPlan, selectedIds),
            {
                onChooseFolder: async (settings) => {
                    const folderPath = await chooseModuleSyncFolder(filePath, settings.folderPath, 'Select export folder');
                    if (!folderPath) {
                        return undefined;
                    }
                    return buildExportSyncPlanFromSettings(filePath, { ...settings, folderPath, folderPathSource: 'session' });
                },
                onRefresh: (settings) => buildExportSyncPlanFromSettings(filePath, settings),
                onReloadWorkbookSettings: () => buildExportSyncPlanFromWorkbookSettings(filePath),
                onSaveSettings: (settings) => saveModuleSyncSettings(filePath, 'xlide.exportModulesToFolder', settings),
            },
        );
        if (!result) {
            return;
        }
        const message = `XLIDE: ${result.summary}`;
        if (result.failed > 0) {
            vscode.window.showWarningMessage(message);
        } else {
            vscode.window.showInformationMessage(message);
        }
    }

    async function showImportModulesDiffGui(filePath: string): Promise<void> {
        const target = await resolveWorkbookImportFolder(filePath);
        if (!target) {
            return;
        }

        log(`[importModules] Workbook: ${filePath}`);
        log(`[importModules] Source folder: ${target.folderPath}`);
        log(`[importModules] Mode: ${target.importMode}`);

        const plan = await buildImportSyncPlanFromSettings(filePath, {
            folderPath: target.folderPath,
            folderPathSource: target.folderPathSource,
            importMode: target.importMode,
            importModeSource: target.importModeSource,
            settingsPath: target.settingsPath,
        });
        const result = await openModuleSyncPreview(
            context,
            plan,
            (currentPlan, selectedIds) => applyImportModuleSyncPlan(currentPlan, selectedIds),
            {
                onChooseFolder: async (settings) => {
                    const folderPath = await chooseModuleSyncFolder(filePath, settings.folderPath, 'Select folder to import from');
                    if (!folderPath) {
                        return undefined;
                    }
                    return buildImportSyncPlanFromSettings(filePath, { ...settings, folderPath, folderPathSource: 'session' });
                },
                onRefresh: (settings) => buildImportSyncPlanFromSettings(filePath, settings),
                onReloadWorkbookSettings: () => buildImportSyncPlanFromWorkbookSettings(filePath),
                onSaveSettings: (settings) => saveModuleSyncSettings(filePath, 'xlide.importModulesFromFolder', settings),
            },
        );
        if (!result) {
            return;
        }
        if (result.failed > 0) {
            vscode.window.showWarningMessage(`XLIDE: ${result.summary}. Copy redacted diagnostics if you need to troubleshoot.`);
        } else {
            vscode.window.showInformationMessage(`XLIDE: ${result.summary} into ${path.basename(filePath)}`);
        }
    }

    async function applyExportModuleSyncPlan(
        plan: ModuleSyncPlan,
        selectedIds: readonly string[],
    ): Promise<ModuleSyncApplyResult> {
        const selected = selectedModuleSyncItems(plan, selectedIds);
        const changed: string[] = [];
        const skipped: string[] = [];
        const removed: string[] = [];
        const failed: string[] = [];

        for (const item of selected) {
            if (item.status === 'unchanged') {
                skipped.push(`${item.relativeName} (unchanged)`);
                continue;
            }
            if (item.status === 'will-remove') {
                try {
                    if (!item.targetPath || !isPathInside(plan.folderPath, item.targetPath)) {
                        throw new Error(`Refusing to remove a file outside the export folder: ${item.relativeName}`);
                    }
                    if (await fileExists(item.targetPath)) {
                        await fs.promises.unlink(item.targetPath);
                        removed.push(item.relativeName);
                    } else {
                        skipped.push(`${item.relativeName} (already missing)`);
                    }
                } catch (err) {
                    failed.push(item.relativeName);
                    log(`[exportModules] Error removing ${item.relativeName}: ${err instanceof Error ? err.message : String(err)}`);
                }
                continue;
            }

            try {
                const result = await exportWorkbookModule(bridge, {
                    filePath: plan.workbookPath,
                    moduleName: item.moduleName,
                    exportFolder: plan.folderPath,
                    exportMode: plan.exportMode,
                });
                changed.push(...result.writtenFiles);
            } catch (err) {
                failed.push(item.relativeName);
                log(`[exportModules] Error exporting ${item.moduleName}: ${err instanceof Error ? err.message : String(err)}`);
            }
        }

        try {
            await persistModuleSyncSettings(plan.workbookPath, syncSettingsFromPlan(plan));
        } catch (err) {
            failed.push('workbook settings');
            recordWriteAudit({
                command: 'xlide.exportModulesToFolder',
                operation: 'configure-module-sync',
                outcome: 'failed',
                workbookPath: plan.workbookPath,
                targetPath: plan.folderPath,
                summary: 'Sync settings: 0 changed, 1 failed',
                error: err,
            });
            log(`[exportModules] Error updating workbook settings: ${err instanceof Error ? err.message : String(err)}`);
        }
        const summaryText = logChangeSummary('exportModules', {
            operation: 'Export modules',
            changed,
            skipped,
            removed,
            failed,
        });
        recordWriteAudit({
            command: 'xlide.exportModulesToFolder',
            operation: 'export-modules',
            outcome: failed.length > 0 ? 'failed' : changed.length > 0 || removed.length > 0 ? 'succeeded' : 'skipped',
            workbookPath: plan.workbookPath,
            targetPath: plan.folderPath,
            summary: summaryText,
        });
        return {
            summary: summaryText,
            changed: changed.length,
            skipped: skipped.length,
            removed: removed.length,
            failed: failed.length,
        };
    }

    async function applyImportModuleSyncPlan(
        plan: ModuleSyncPlan,
        selectedIds: readonly string[],
    ): Promise<ModuleSyncApplyResult> {
        const selected = selectedModuleSyncItems(plan, selectedIds);
        const changed: string[] = [];
        const skipped: string[] = [];
        const removed: string[] = [];
        const failed: string[] = [];

        for (const item of selected) {
            if (item.status === 'unchanged') {
                skipped.push(`${item.relativeName} (unchanged)`);
                continue;
            }
            if (item.status === 'will-remove') {
                try {
                    log(`[importModules] Deleting workbook module ${item.moduleName} during import true-up`);
                    const result = await bridge.call<{ ok?: boolean; signatureDropped?: boolean }>('deleteModule', {
                        path: plan.workbookPath,
                        module: item.moduleName,
                    });
                    notifySignatureDropped(plan.workbookPath, Boolean(result.signatureDropped));
                    removed.push(item.relativeName);
                    recordWriteAudit({
                        command: 'xlide.importModulesFromFolder',
                        operation: 'delete-module',
                        outcome: 'succeeded',
                        workbookPath: plan.workbookPath,
                        moduleName: item.moduleName,
                        summary: 'Import true-up: 1 removed',
                    });
                } catch (err) {
                    failed.push(item.relativeName);
                    recordWriteAudit({
                        command: 'xlide.importModulesFromFolder',
                        operation: 'delete-module',
                        outcome: 'failed',
                        workbookPath: plan.workbookPath,
                        moduleName: item.moduleName,
                        summary: 'Import true-up: 0 removed, 1 failed',
                        error: err,
                    });
                    log(`[importModules] Error deleting ${item.moduleName}: ${err instanceof Error ? err.message : String(err)}`);
                }
                continue;
            }
            if (item.status === 'skipping-import' || (item.unsupportedDirectCreation && !item.existsInWorkbook)) {
                skipped.push(`${item.relativeName} (${item.moduleType} cannot be created directly)`);
                recordWriteAudit({
                    command: 'xlide.importModulesFromFolder',
                    operation: 'import-module',
                    outcome: 'skipped',
                    workbookPath: plan.workbookPath,
                    moduleName: item.moduleName,
                    sourcePath: item.sourcePath,
                    summary: 'Import module: 0 changed, 1 skipped',
                });
                continue;
            }

            try {
                if (!item.sourcePath) {
                    throw new Error(`Missing source path for ${item.moduleName}.`);
                }
                const source = await fs.promises.readFile(item.sourcePath, 'utf8');
                log(`[importModules] Importing ${item.moduleName} from ${item.relativeName}`);
                const result = await bridge.call<{ ok?: boolean; signatureDropped?: boolean }>('writeModule', {
                    path: plan.workbookPath,
                    module: item.moduleName,
                    source,
                    kind: item.moduleType,
                });
                notifySignatureDropped(plan.workbookPath, Boolean(result.signatureDropped));
                changed.push(item.relativeName);
                recordWriteAudit({
                    command: 'xlide.importModulesFromFolder',
                    operation: 'import-module',
                    outcome: 'succeeded',
                    workbookPath: plan.workbookPath,
                    moduleName: item.moduleName,
                    sourcePath: item.sourcePath,
                    summary: 'Import module: 1 changed',
                });
            } catch (err) {
                failed.push(item.relativeName);
                recordWriteAudit({
                    command: 'xlide.importModulesFromFolder',
                    operation: 'import-module',
                    outcome: 'failed',
                    workbookPath: plan.workbookPath,
                    moduleName: item.moduleName,
                    sourcePath: item.sourcePath,
                    summary: 'Import module: 0 changed, 1 failed',
                    error: err,
                });
                log(`[importModules] Error importing ${item.moduleName}: ${err instanceof Error ? err.message : String(err)}`);
            }
        }

        if (changed.length > 0 || removed.length > 0) {
            refreshVbaProjectState(plan.workbookPath);
        }
        try {
            await persistModuleSyncSettings(plan.workbookPath, syncSettingsFromPlan(plan));
        } catch (err) {
            failed.push('workbook settings');
            recordWriteAudit({
                command: 'xlide.importModulesFromFolder',
                operation: 'configure-module-sync',
                outcome: 'failed',
                workbookPath: plan.workbookPath,
                targetPath: plan.folderPath,
                summary: 'Sync settings: 0 changed, 1 failed',
                error: err,
            });
        }
        const summaryText = logChangeSummary('importModules', {
            operation: 'Import modules',
            changed,
            skipped,
            removed,
            failed,
        });
        return {
            summary: summaryText,
            changed: changed.length,
            skipped: skipped.length,
            removed: removed.length,
            failed: failed.length,
        };
    }

    function selectedModuleSyncItems(
        plan: ModuleSyncPlan,
        selectedIds: readonly string[],
    ): ModuleSyncPlanItem[] {
        const selected = new Set(selectedIds);
        return plan.items.filter((item) => selected.has(item.id));
    }

    function isPathInside(baseDir: string, targetPath: string): boolean {
        const base = path.resolve(baseDir);
        const target = path.resolve(targetPath);
        return target === base || target.startsWith(base + path.sep);
    }

    async function fileExists(filePath: string): Promise<boolean> {
        try {
            await fs.promises.access(filePath, fs.constants.F_OK);
            return true;
        } catch {
            return false;
        }
    }

    async function currentModuleAnalysisResult(document: vscode.TextDocument): Promise<WorkbookAnalysisResult> {
        const { xlsmPath, moduleName } = decodeModuleUri(document.uri);
        const source = document.getText();
        const modules = applyOpenDocumentSources(
            await vbaIndex.getAllModules(xlsmPath),
            xlsmPath,
        );
        const current = modules.find(
            (mod) => mod.moduleName.toLowerCase() === moduleName.toLowerCase(),
        );
        const moduleKind = moduleKindFromType(current?.type);
        const moduleType = current?.type ?? 'standard';
        const project = buildLiveVbaProjectIndex(modules, {
            moduleName,
            moduleKind,
            source,
        });
        const projectOptions = projectAnalysisOptionsForModule(
            project,
            moduleName,
            projectProcedureSignatures(project),
        );
        const analysisSettings = await effectiveWorkbookAnalysisSettings(xlsmPath);

        const result = analyzeVbaModuleSource({
            source,
            moduleName,
            moduleType,
            moduleKind,
            documentType: current?.documentType,
            severityOverrides: analysisSettings.ruleSeverityOverrides,
            ...projectOptions,
        });
        const problems = workbookProblemsForModule(
            moduleName,
            moduleType,
            source,
            result.diagnostics,
        ).sort((a, b) => {
            if (a.line !== b.line) { return a.line - b.line; }
            return a.column - b.column;
        });
        const suppressedProblems = workbookProblemsForModule(
            moduleName,
            moduleType,
            source,
            result.suppressedDiagnostics,
            { suppressed: true },
        ).sort((a, b) => {
            if (a.line !== b.line) { return a.line - b.line; }
            return a.column - b.column;
        });
        const errorCount = problems.filter((p) => p.severity === 'error').length;
        const warningCount = problems.filter((p) => p.severity === 'warning').length;
        const summary = summarizeWorkbookAnalysisProblems(problems, suppressedProblems.length);
        return {
            filePath: xlsmPath,
            moduleCount: 1,
            problems,
            suppressedProblems,
            errorCount,
            warningCount,
            summary,
        };
    }

    async function analyzeActiveModule(): Promise<void> {
        const editor = vscode.window.activeTextEditor;
        if (!editor || editor.document.uri.scheme !== XLIDE_SCHEME) {
            vscode.window.showWarningMessage('XLIDE: Open a workbook VBA module to analyze the current module.');
            return;
        }

        const { moduleName } = decodeModuleUri(editor.document.uri);
        const analysisResult = await currentModuleAnalysisResult(editor.document);
        const { problems, errorCount, warningCount } = analysisResult;

        showWorkbookAnalysisResults(analysisResult, () => currentModuleAnalysisResult(editor.document));
        if (problems.length === 0) {
            vscode.window.showInformationMessage(
                `XLIDE: "${moduleName}" passed analysis (no unsuppressed problems).`,
            );
        } else {
            vscode.window.showWarningMessage(
                `XLIDE: "${moduleName}" has ${errorCount} error(s) and ${warningCount} warning(s).`,
            );
        }
    }

    interface SupportBundleOptions {
        includeAnonymizedWorkbookAnalysisReport?: boolean;
        includeSelectedLogs?: boolean;
    }

    async function activeLocalWorkbookPath(): Promise<string | undefined> {
        const editor = vscode.window.activeTextEditor;
        if (!editor || editor.document.uri.scheme !== XLIDE_SCHEME || editor.document.uri.authority) {
            return undefined;
        }
        return decodeModuleUri(editor.document.uri).xlsmPath;
    }

    async function activeModuleSupportData(): Promise<{
        workbook: SupportBundleWorkbookSummary;
        analysis: SupportBundleAnalysisSummary;
    }> {
        const editor = vscode.window.activeTextEditor;
        if (!editor || editor.document.uri.scheme !== XLIDE_SCHEME || editor.document.uri.authority) {
            return {
                workbook: { available: false },
                analysis: { available: false },
            };
        }

        const { xlsmPath, moduleName } = decodeModuleUri(editor.document.uri);
        const modules = applyOpenDocumentSources(
            await vbaIndex.getAllModules(xlsmPath),
            xlsmPath,
        );
        const current = modules.find(
            (mod) => mod.moduleName.toLowerCase() === moduleName.toLowerCase(),
        );
        const moduleType = current?.type ?? 'standard';
        const moduleTypes = countBy(modules.map((mod) => mod.type || 'unknown'));
        const workbook: SupportBundleWorkbookSummary = {
            available: true,
            workbookPath: xlsmPath,
            extension: path.extname(xlsmPath).toLowerCase(),
            moduleCount: modules.length,
            moduleTypes,
            activeModuleType: moduleType,
        };

        const source = editor.document.getText();
        const moduleKind = moduleKindFromType(current?.type);
        const project = buildLiveVbaProjectIndex(modules, {
            moduleName,
            moduleKind,
            source,
        });
        const projectOptions = projectAnalysisOptionsForModule(
            project,
            moduleName,
            projectProcedureSignatures(project),
        );
        const analysisSettings = await effectiveWorkbookAnalysisSettings(xlsmPath);
        const result = analyzeVbaModuleSource({
            source,
            moduleName,
            moduleType,
            moduleKind,
            documentType: current?.documentType,
            severityOverrides: analysisSettings.ruleSeverityOverrides,
            ...projectOptions,
        });
        const problems = workbookProblemsForModule(
            moduleName,
            moduleType,
            source,
            result.diagnostics,
        );
        return {
            workbook,
            analysis: {
                available: true,
                moduleType,
                errorCount: problems.filter((problem) => problem.severity === 'error').length,
                warningCount: problems.filter((problem) => problem.severity === 'warning').length,
                suppressedCount: result.suppressedCount,
                byCode: countBy(problems.map((problem) => problem.code || 'unclassified')),
            },
        };
    }

    function countBy(values: readonly string[]): Record<string, number> {
        const out: Record<string, number> = {};
        for (const value of values) {
            out[value] = (out[value] ?? 0) + 1;
        }
        return out;
    }

    function xlideSettingsForSupportBundle(): SupportBundleSetting[] {
        return resolvedXlideGlobalSettingsFromConfig(vscode.workspace.getConfiguration('xlide'));
    }

    async function anonymizedWorkbookAnalysisReportForActiveWorkbook():
        Promise<SupportBundleAnonymizedAnalysisReport> {
        const workbookPath = await activeLocalWorkbookPath();
        if (!workbookPath) {
            return { included: false, unavailableReason: 'no-active-workbook' };
        }
        try {
            return anonymizedWorkbookAnalysisReportFromResult(await analyzeWorkbook(bridge, workbookPath));
        } catch (err) {
            return {
                included: false,
                unavailableReason: 'analysis-failed',
                errorCategory: errorCategoryForSupportLog(err),
            };
        }
    }

    async function currentSupportBundle(
        now = new Date(),
        options: SupportBundleOptions = {},
    ): Promise<SupportBundle> {
        const packageJson = context.extension.packageJSON as {
            name?: string;
            publisher?: string;
            version?: string;
            displayName?: string;
        };
        const active = await activeModuleSupportData();
        const anonymizedWorkbookAnalysisReport = options.includeAnonymizedWorkbookAnalysisReport
            ? await anonymizedWorkbookAnalysisReportForActiveWorkbook()
            : undefined;
        return buildSupportBundle({
            generatedAt: now.toISOString(),
            extension: {
                id: packageJson.publisher && packageJson.name
                    ? `${packageJson.publisher}.${packageJson.name}`
                    : packageJson.name,
                name: packageJson.displayName ?? packageJson.name,
                version: packageJson.version,
            },
            vscode: {
                version: vscode.version,
                appName: vscode.env.appName,
            },
            runtime: {
                platform: process.platform,
                arch: process.arch,
                node: process.version,
            },
            workspace: {
                folderCount: vscode.workspace.workspaceFolders?.length ?? 0,
            },
            settings: xlideSettingsForSupportBundle(),
            workbook: active.workbook,
            analysis: active.analysis,
            commands: recentXlideCommands(),
            writeAudits: recentXlideWriteAudits(),
            anonymizedWorkbookAnalysisReport,
            selectedLogs: options.includeSelectedLogs ? recentXlideOutputLog() : undefined,
        });
    }

    async function selectSupportBundleExportOptions(
        bundle: SupportBundle,
    ): Promise<SupportBundleOptions | undefined> {
        const choice = await vscode.window.showInformationMessage(
            'XLIDE support bundle export',
            {
                modal: true,
                detail: supportBundleDisclosureText(bundle),
            },
            'Export',
            'Choose Extras',
            'Copy Diagnostics',
        );
        if (choice === 'Copy Diagnostics') {
            await copyDiagnosticsFromBundle(bundle);
            return undefined;
        }
        if (choice === 'Choose Extras') {
            const picks = await vscode.window.showQuickPick(
                [
                    {
                        label: 'Anonymized workbook analysis report',
                        description: 'Counts by rule/module type only; no source or module names',
                        picked: true,
                        option: 'includeAnonymizedWorkbookAnalysisReport' as const,
                    },
                    {
                        label: 'Selected recent XLIDE logs',
                        description: 'Recent XLIDE-authored output lines with paths redacted',
                        picked: false,
                        option: 'includeSelectedLogs' as const,
                    },
                ],
                {
                    title: 'XLIDE Support Bundle: Optional Extras',
                    canPickMany: true,
                    placeHolder: 'Choose only the extras you want included in this export.',
                },
            );
            if (!picks) {
                return undefined;
            }
            return {
                includeAnonymizedWorkbookAnalysisReport:
                    picks.some((pick) => pick.option === 'includeAnonymizedWorkbookAnalysisReport'),
                includeSelectedLogs:
                    picks.some((pick) => pick.option === 'includeSelectedLogs'),
            };
        }
        return choice === 'Export' ? {} : undefined;
    }

    async function copyDiagnosticsFromBundle(bundle: SupportBundle): Promise<void> {
        await vscode.env.clipboard.writeText(supportDiagnosticsText(bundle));
        vscode.window.showInformationMessage('XLIDE: Redacted diagnostics copied to clipboard.');
    }

    async function copyDiagnostics(): Promise<void> {
        await copyDiagnosticsFromBundle(await currentSupportBundle());
    }

    async function exportSupportBundle(): Promise<void> {
        const now = new Date();
        const baseBundle = await currentSupportBundle(now);
        const options = await selectSupportBundleExportOptions(baseBundle);
        if (!options) {
            return;
        }
        const bundle = Object.keys(options).length === 0
            ? baseBundle
            : await currentSupportBundle(now, options);

        const defaultFolder = vscode.workspace.workspaceFolders?.[0]?.uri;
        const target = await vscode.window.showSaveDialog({
            title: 'Export XLIDE Support Bundle (redacted JSON; no workbook source)',
            defaultUri: defaultFolder
                ? vscode.Uri.joinPath(defaultFolder, defaultSupportBundleFileName(now))
                : undefined,
            filters: { JSON: ['json'] },
        });
        if (!target) {
            return;
        }

        await fs.promises.writeFile(target.fsPath, `${JSON.stringify(bundle, null, 2)}\n`, 'utf8');
        vscode.window.showInformationMessage(`XLIDE: Support bundle exported to ${target.fsPath}`);
    }

    async function showClassModuleReferences(node: XlideNode): Promise<void> {
        if (!node.moduleName || !node.filePath || node.isRemote) { return; }
        const originUri = encodeModuleUri(node.filePath, node.moduleName);
        const originDoc = await vscode.workspace.openTextDocument(originUri);
        await vscode.languages.setTextDocumentLanguage(originDoc, 'vba');
        const editor = await vscode.window.showTextDocument(originDoc, { preview: false });
        const origin = new vscode.Position(0, 0);
        editor.selection = new vscode.Selection(origin, origin);
        await vscode.commands.executeCommand('references-view.findReferences', originUri, origin);
    }

    return [
        registerXlideCommand('xlide.refreshExplorer', () => {
            explorer.refresh();
        }),

        // Open a module (or navigate to a sub's line inside one)
        registerXlideCommand('xlide.openModule', async (node: XlideNode) => {
            if (!node?.moduleName) { return; }
            const uri = node.isRemote && node.remoteId
                ? encodeRemoteModuleUri(node.remoteId, node.moduleName)
                : encodeModuleUri(node.filePath, node.moduleName);

            // Set the language to 'vba' so syntax highlighters kick in
            const doc = await vscode.workspace.openTextDocument(uri);
            const editor = await vscode.window.showTextDocument(doc, { preview: false });

            // If a specific line was requested (sub navigation), move cursor there
            if (node.line !== undefined && node.line > 0) {
                const pos = new vscode.Position(node.line - 1, 0);
                editor.selection = new vscode.Selection(pos, pos);
                editor.revealRange(
                    new vscode.Range(pos, pos),
                    vscode.TextEditorRevealType.InCenterIfOutsideViewport,
                );
            }

            // Set language mode to vba for the document
            await vscode.languages.setTextDocumentLanguage(doc, 'vba');
        }),

        // Find all references to the procedure or class represented by a tree node
        registerXlideCommand('xlide.findReferences', async (node: XlideNode) => {
            if (!node?.moduleName) { return; }
            if (node.kind === 'module' && node.moduleType === 'class') {
                await showClassModuleReferences(node);
                return;
            }
            if (node.kind !== 'sub') { return; }
            const uri = node.isRemote && node.remoteId
                ? encodeRemoteModuleUri(node.remoteId, node.moduleName)
                : encodeModuleUri(node.filePath, node.moduleName);

            const doc = await vscode.workspace.openTextDocument(uri);
            await vscode.languages.setTextDocumentLanguage(doc, 'vba');
            const editor = await vscode.window.showTextDocument(doc, { preview: false });

            // Locate the procedure name on its declaration line so the reference
            // search starts on the identifier. The node label is "<kind> <name>"
            // (kind may be "Property Get" etc.), so the bare name is the last token.
            const procName = node.label.split(' ').pop() ?? '';
            let pos = new vscode.Position(Math.max(0, (node.line ?? 1) - 1), 0);
            if (procName && node.line !== undefined && node.line > 0) {
                const lineText = doc.lineAt(node.line - 1).text;
                const col = lineText.indexOf(procName);
                if (col >= 0) {
                    pos = new vscode.Position(node.line - 1, col);
                }
            }

            // Move the active editor's cursor onto the identifier so the
            // references command resolves the correct symbol, then trigger it.
            editor.selection = new vscode.Selection(pos, pos);
            editor.revealRange(
                new vscode.Range(pos, pos),
                vscode.TextEditorRevealType.InCenterIfOutsideViewport,
            );
            await vscode.commands.executeCommand('references-view.findReferences', uri, pos);
        }),

        registerXlideCommand('xlide.newModule', async (node: XlideNode) => {
            if (node?.kind !== 'xlsm') { return; }
            const name = await vscode.window.showInputBox({
                prompt: 'New module name',
                placeHolder: 'Module1',
                validateInput: (v) =>
                    /^\w+$/.test(v) ? undefined : 'Module names must be alphanumeric',
            });
            if (!name) { return; }

            const stub = `Option Explicit\r\n\r\nSub ${name}_Main()\r\n\r\nEnd Sub\r\n`;
            try {
                await bridge.call('writeModule', {
                    path: node.filePath,
                    module: name,
                    source: stub,
                });
                const summaryText = logChangeSummary('newModule', {
                    operation: 'Create module',
                    changed: [name],
                });
                recordWriteAudit({
                    command: 'xlide.newModule',
                    operation: 'create-module',
                    outcome: 'succeeded',
                    workbookPath: node.filePath,
                    moduleName: name,
                    summary: summaryText,
                });
                refreshVbaProjectState(node.filePath);
                // Open the new module immediately
                const uri = encodeModuleUri(node.filePath, name);
                const doc = await vscode.workspace.openTextDocument(uri);
                await vscode.window.showTextDocument(doc, { preview: false });
                await vscode.languages.setTextDocumentLanguage(doc, 'vba');
            } catch (err) {
                recordWriteAudit({
                    command: 'xlide.newModule',
                    operation: 'create-module',
                    outcome: 'failed',
                    workbookPath: node.filePath,
                    moduleName: name,
                    summary: 'Create module: 0 changed, 1 failed',
                    error: err,
                });
                vscode.window.showErrorMessage(`XLIDE: Failed to create module: ${err}`);
            }
        }),

        // Add a new class module
        registerXlideCommand('xlide.newClassModule', async (node: XlideNode) => {
            if (node?.kind !== 'xlsm') { return; }
            const name = await vscode.window.showInputBox({
                prompt: 'New class module name',
                placeHolder: 'MyClass',
                validateInput: (v) =>
                    VBA_IDENTIFIER_NAME_RE.test(v) ? undefined : 'Class module names must be valid VBA identifiers',
            });
            if (!name) { return; }

            const stub = `Option Explicit\r\n\r\nPrivate Sub Class_Initialize()\r\n\r\nEnd Sub\r\n\r\nPrivate Sub Class_Terminate()\r\n\r\nEnd Sub\r\n`;
            try {
                await bridge.call('writeModule', {
                    path: node.filePath,
                    module: name,
                    source: stub,
                    kind: 'class',
                });
                const summaryText = logChangeSummary('newClassModule', {
                    operation: 'Create class module',
                    changed: [name],
                });
                recordWriteAudit({
                    command: 'xlide.newClassModule',
                    operation: 'create-class-module',
                    outcome: 'succeeded',
                    workbookPath: node.filePath,
                    moduleName: name,
                    summary: summaryText,
                });
                refreshVbaProjectState(node.filePath);
                const uri = encodeModuleUri(node.filePath, name);
                const doc = await vscode.workspace.openTextDocument(uri);
                await vscode.window.showTextDocument(doc, { preview: false });
                await vscode.languages.setTextDocumentLanguage(doc, 'vba');
            } catch (err) {
                recordWriteAudit({
                    command: 'xlide.newClassModule',
                    operation: 'create-class-module',
                    outcome: 'failed',
                    workbookPath: node.filePath,
                    moduleName: name,
                    summary: 'Create class module: 0 changed, 1 failed',
                    error: err,
                });
                vscode.window.showErrorMessage(`XLIDE: Failed to create class module: ${err}`);
            }
        }),

        // Rename a module
        registerXlideCommand('xlide.renameModule', async (node: XlideNode) => {
            if (!node?.moduleName) { return; }
            const validateInput = (v: string): string | undefined => {
                if (node.moduleType === 'class') {
                    return VBA_IDENTIFIER_NAME_RE.test(v)
                        ? undefined
                        : 'Class module names must be valid VBA identifiers';
                }
                return /^\w+$/.test(v) ? undefined : 'Module names must be alphanumeric';
            };
            const newName = await vscode.window.showInputBox({
                prompt: `Rename "${node.moduleName}" to`,
                value: node.moduleName,
                validateInput,
            });
            if (!newName || newName === node.moduleName) { return; }

            let moduleRenamed = false;
            try {
                if (node.moduleType === 'class') {
                    const modules = applyOpenDocumentSources(
                        await vbaIndex.getAllModules(node.filePath),
                        node.filePath,
                    );
                    const project = buildVbaProjectIndex(modules);
                    const byModule = new Map(modules.map((mod) => [mod.moduleName.toLowerCase(), mod]));
                    const definition = projectClassModuleDefinition(
                        project,
                        node.moduleName,
                        node.moduleName,
                    );
                    if (!definition) {
                        throw new Error(`"${node.moduleName}" is not a project-defined class module.`);
                    }
                    const references = projectClassReferenceEdit(
                        node.filePath,
                        byModule,
                        project,
                        node.moduleName,
                        definition,
                        newName,
                    );
                    await renameProjectClassModule(bridge, node.filePath, node.moduleName, newName);
                    moduleRenamed = true;
                    vbaIndex.invalidate(node.filePath);
                    if (references.count > 0) {
                        for (const uri of references.uris) {
                            const doc = await vscode.workspace.openTextDocument(uri);
                            await vscode.languages.setTextDocumentLanguage(doc, 'vba');
                        }
                        const applied = await vscode.workspace.applyEdit(references.edit);
                        if (!applied) {
                            throw new Error('VS Code did not apply the class reference edits.');
                        }
                    }
                } else {
                    const result = await bridge.call<{ ok: boolean; signatureDropped: boolean }>(
                        'renameModule',
                        {
                            path: node.filePath,
                            module: node.moduleName,
                            newName,
                        },
                    );
                    moduleRenamed = true;
                    notifySignatureDropped(node.filePath, result.signatureDropped);
                    vbaIndex.invalidate(node.filePath);
                }
                const summaryText = logChangeSummary('renameModule', {
                    operation: 'Rename module',
                    changed: [`${node.moduleName} -> ${newName}`],
                });
                recordWriteAudit({
                    command: 'xlide.renameModule',
                    operation: 'rename-module',
                    outcome: 'succeeded',
                    workbookPath: node.filePath,
                    moduleName: newName,
                    summary: summaryText,
                });
            } catch (err) {
                const prefix = moduleRenamed
                    ? 'XLIDE: Module was renamed, but reference updates failed'
                    : 'XLIDE: Rename failed';
                recordWriteAudit({
                    command: 'xlide.renameModule',
                    operation: 'rename-module',
                    outcome: 'failed',
                    workbookPath: node.filePath,
                    moduleName: moduleRenamed ? newName : node.moduleName,
                    summary: moduleRenamed
                        ? 'Rename module: 1 changed, 1 failed'
                        : 'Rename module: 0 changed, 1 failed',
                    error: err,
                });
                vscode.window.showErrorMessage(`${prefix}: ${err}`);
            } finally {
                if (moduleRenamed) {
                    refreshVbaProjectState(node.filePath);
                }
            }
        }),

        // Delete a module (with confirmation)
        registerXlideCommand('xlide.deleteModule', async (node: XlideNode) => {
            if (!node?.moduleName) { return; }

            // Prevent deletion of document-type modules
            if (node.moduleType === 'document') {
                vscode.window.showWarningMessage(
                    `Cannot delete "${node.moduleName}" — document modules are protected.`,
                );
                return;
            }

            const choice = await vscode.window.showWarningMessage(
                `Delete module "${node.moduleName}" from "${path.basename(node.filePath)}"?`,
                { modal: true },
                'Delete',
            );
            if (choice !== 'Delete') { return; }

            try {
                const result = await bridge.call<{ ok: boolean; signatureDropped: boolean }>(
                    'deleteModule',
                    {
                        path: node.filePath,
                        module: node.moduleName,
                    },
                );
                notifySignatureDropped(node.filePath, result.signatureDropped);
                const summaryText = logChangeSummary('deleteModule', {
                    operation: 'Delete module',
                    changed: [node.moduleName],
                });
                recordWriteAudit({
                    command: 'xlide.deleteModule',
                    operation: 'delete-module',
                    outcome: 'succeeded',
                    workbookPath: node.filePath,
                    moduleName: node.moduleName,
                    summary: summaryText,
                });
                // Close any open editors for this module
                const uri = encodeModuleUri(node.filePath, node.moduleName);
                for (const tab of vscode.window.tabGroups.all.flatMap((g) => g.tabs)) {
                    const input = tab.input;
                    if (
                        input instanceof vscode.TabInputText &&
                        input.uri.toString() === uri.toString()
                    ) {
                        await vscode.window.tabGroups.close(tab);
                    }
                }
                refreshVbaProjectState(node.filePath);
            } catch (err) {
                recordWriteAudit({
                    command: 'xlide.deleteModule',
                    operation: 'delete-module',
                    outcome: 'failed',
                    workbookPath: node.filePath,
                    moduleName: node.moduleName,
                    summary: 'Delete module: 0 changed, 1 failed',
                    error: err,
                });
                vscode.window.showErrorMessage(`XLIDE: Delete failed: ${err}`);
            }
        }),

        // Export all modules to a user-selected folder and persist folder in workbook config JSON
        registerXlideCommand('xlide.exportModulesToFolder', async (node: XlideNode) => {
            const filePath = resolveWorkbookPath(node);
            if (!filePath) { return; }

            try {
                await showExportModulesDiffGui(filePath);
            } catch (err) {
                const message = err instanceof Error ? err.message : String(err);
                log(`[exportModules] Error: ${message}`);
                recordWriteAudit({
                    command: 'xlide.exportModulesToFolder',
                    operation: 'export-modules',
                    outcome: 'failed',
                    workbookPath: filePath,
                    summary: 'Export modules: 0 changed, 1 failed',
                    error: err,
                });
                vscode.window.showErrorMessage(`XLIDE: Failed to export modules: ${message}`);
            }
        }),

        // Save and export just the active VBA module to the configured module folder
        registerXlideCommand('xlide.exportCurrentModuleToFolder', async () => {
            try {
                await exportActiveModule();
            } catch (err) {
                const message = err instanceof Error ? err.message : String(err);
                log(`[exportCurrentModule] Error: ${message}`);
                recordWriteAudit({
                    command: 'xlide.exportCurrentModuleToFolder',
                    operation: 'export-current-module',
                    outcome: 'failed',
                    workbookPath: await activeLocalWorkbookPath(),
                    summary: 'Export current module: 0 changed, 1 failed',
                    error: err,
                });
                vscode.window.showErrorMessage(`XLIDE: Failed to export current module: ${message}`);
            }
        }),

        // Import selected module files from the configured (or user-chosen) export folder
        registerXlideCommand('xlide.importModulesFromFolder', async (node: XlideNode) => {
            const filePath = resolveWorkbookPath(node);
            if (!filePath) { return; }

            try {
                await showImportModulesDiffGui(filePath);
            } catch (err) {
                const message = err instanceof Error ? err.message : String(err);
                log(`[importModules] Error: ${message}`);
                vscode.window.showErrorMessage(`XLIDE: Import failed: ${message}`);
            }
        }),

        // Export a redacted local diagnostic snapshot for support/self-debugging.
        registerXlideCommand('xlide.exportSupportBundle', async () => {
            try {
                await exportSupportBundle();
            } catch (err) {
                const message = err instanceof Error ? err.message : String(err);
                log(`[supportBundle] Error: ${message}`);
                vscode.window.showErrorMessage(`XLIDE: Failed to export support bundle: ${message}`);
            }
        }),

        registerXlideCommand('xlide.copyDiagnostics', async () => {
            try {
                await copyDiagnostics();
            } catch (err) {
                const message = err instanceof Error ? err.message : String(err);
                log(`[copyDiagnostics] Error: ${message}`);
                vscode.window.showErrorMessage(`XLIDE: Failed to copy diagnostics: ${message}`);
            }
        }),

        // DEV: smoke test — verifies listModules + readModule against a workspace workbook
        registerXlideCommand('xlide.dev.smoke', async () => {
            log('[smoke] Starting smoke test...');

            const uris = (await vscode.workspace.findFiles('**/*.{xlsm,xlsb,xlam}',
                '{**/node_modules/**,**/.venv/**,**/venv/**}'))
                .filter(u => !path.basename(u.fsPath).startsWith('~$'));

            if (uris.length === 0) {
                vscode.window.showErrorMessage('XLIDE Smoke: No workbook found in the workspace.');
                return;
            }

            let workbookPath: string;
            if (uris.length === 1) {
                workbookPath = uris[0].fsPath;
            } else {
                const pick = await vscode.window.showQuickPick(
                    uris.map(u => ({ label: path.basename(u.fsPath), description: u.fsPath, fsPath: u.fsPath })),
                    { title: 'XLIDE Smoke Test: pick a workbook' },
                );
                if (!pick) { return; }
                workbookPath = pick.fsPath;
            }

            log(`[smoke] Workbook: ${workbookPath}`);

            await vscode.window.withProgress(
                { location: vscode.ProgressLocation.Notification, title: 'XLIDE: Running smoke test...', cancellable: false },
                async () => {
                    try {
                        // Step 1: listModules
                        const modules = await bridge.call<Array<{ name: string; type: string }>>(
                            'listModules', { path: workbookPath },
                        );
                        log(`[smoke] listModules OK — ${modules.length} module(s): ${modules.map(m => m.name).join(', ')}`);

                        if (modules.length === 0) {
                            vscode.window.showWarningMessage('XLIDE Smoke: workbook has no VBA modules.');
                            return;
                        }

                        // Step 2: readModule (prefer a non-document module)
                        const target = modules.find(m => m.type !== 'document') ?? modules[0];
                        const source = await bridge.call<string>(
                            'readModule', { path: workbookPath, module: target.name, full: false },
                        );
                        log(`[smoke] readModule "${target.name}" OK — ${source.length} chars`);

                        log('[smoke] All checks passed.');
                        void vscode.window.showInformationMessage(
                            `XLIDE Smoke: OK — ${modules.length} modules, read "${target.name}" (${source.length} chars). See XLIDE Output for details.`,
                        );
                    } catch (err) {
                        const msg = err instanceof Error ? err.message : String(err);
                        log(`[smoke] FAILED: ${msg}`);
                        vscode.window.showErrorMessage(`XLIDE Smoke FAILED: ${msg}`);
                    }
                },
            );
        }),

        // Validate the workbook's VBA project structure
        registerXlideCommand('xlide.validateWorkbook', async (node: XlideNode) => {
            const filePath = resolveWorkbookPath(node);
            if (!filePath) {
                vscode.window.showWarningMessage('XLIDE: No workbook selected to validate.');
                return;
            }
            const name = path.basename(filePath);
            await vscode.window.withProgress(
                { location: vscode.ProgressLocation.Notification, title: `XLIDE: Validating "${name}"...`, cancellable: false },
                async () => {
                    try {
                        const res = await bridge.call<{ issues: string[] }>('validateWorkbook', { path: filePath });
                        const issues = res.issues ?? [];
                        if (issues.length === 0) {
                            log(`[validate] "${name}": no issues`);
                            void vscode.window.showInformationMessage(`XLIDE: "${name}" passed validation (no issues).`);
                            return;
                        }
                        log(`[validate] "${name}": ${issues.length} issue(s):`);
                        for (const issue of issues) {
                            log(`[validate]   - ${issue}`);
                        }
                        void vscode.window.showWarningMessage(
                            `XLIDE: "${name}" has ${issues.length} validation issue(s). See XLIDE Output for details.`,
                        );
                    } catch (err) {
                        const msg = err instanceof Error ? err.message : String(err);
                        log(`[validate] FAILED: ${msg}`);
                        vscode.window.showErrorMessage(`XLIDE: Validation failed: ${msg}`);
                    }
                },
            );
        }),

        // Analyze the active VBA module using the same source text the editor shows.
        registerXlideCommand('xlide.analyzeCurrentModule', async () => {
            try {
                await analyzeActiveModule();
            } catch (err) {
                const message = err instanceof Error ? err.message : String(err);
                log(`[analyzeCurrentModule] Error: ${message}`);
                vscode.window.showErrorMessage(`XLIDE: Failed to analyze current module: ${message}`);
            }
        }),

        // Analyze every VBA module in the workbook and show a navigable results panel.
        registerXlideCommand('xlide.analyzeWorkbook', async (node: XlideNode) => {
            const filePath = resolveWorkbookPath(node);
            if (!filePath) {
                vscode.window.showWarningMessage('XLIDE: No workbook selected to analyze.');
                return;
            }
            const name = path.basename(filePath);
            await vscode.window.withProgress(
                { location: vscode.ProgressLocation.Notification, title: `XLIDE: Analyzing "${name}"...`, cancellable: false },
                async () => {
                    try {
                        const result = await analyzeWorkbook(bridge, filePath);
                        showWorkbookAnalysisResults(result, () => analyzeWorkbook(bridge, filePath));
                        if (result.problems.length === 0) {
                            vscode.window.showInformationMessage(
                                `XLIDE: "${name}" passed analysis (no problems across ${result.moduleCount} module(s)).`,
                            );
                        } else {
                            vscode.window.showWarningMessage(
                                `XLIDE: "${name}" has ${result.errorCount} error(s) and ${result.warningCount} warning(s).`,
                            );
                        }
                    } catch (err) {
                        const msg = err instanceof Error ? err.message : String(err);
                        log(`[analyzeWorkbook] FAILED: ${msg}`);
                        vscode.window.showErrorMessage(`XLIDE: Analysis failed: ${msg}`);
                    }
                },
            );
        }),

        // Open the workbook-scoped VBA tests GUI.
        registerXlideCommand('xlide.runVbaTests', async (node: XlideNode) => {
            const filePath = resolveWorkbookPath(node);
            if (!filePath) {
                vscode.window.showWarningMessage('XLIDE: No workbook selected to test.');
                return;
            }
            openVbaTestsForWorkbook(filePath);
        }),

        // Run workbook VBA tests with tag include/exclude filters and fail-fast mode.
        registerXlideCommand('xlide.runVbaTestsWithFilters', async (node: XlideNode) => {
            const filePath = resolveWorkbookPath(node);
            if (!filePath) {
                vscode.window.showWarningMessage('XLIDE: No workbook selected to test.');
                return;
            }
            const options = await promptVbaTestRunOptions();
            if (!options) {
                return;
            }
            await runVbaTestsForWorkbook(filePath, options);
        }),

        // Run all annotated tests in the active VBA module.
        registerXlideCommand('xlide.runVbaTestsInCurrentModule', async () => {
            await runCurrentModuleVbaTests();
        }),

        // Run the annotated test procedure under the active editor cursor.
        registerXlideCommand('xlide.runVbaTestAtCursor', async () => {
            await runVbaTestAtCursor();
        }),

        // Install/update the standard VBA assertion helpers used by XLIDE tests.
        registerXlideCommand('xlide.installVbaTestSupport', async (node: XlideNode) => {
            const filePath = resolveWorkbookPath(node);
            if (!filePath) {
                vscode.window.showWarningMessage('XLIDE: No workbook selected for test support installation.');
                return;
            }
            try {
                await installVbaTestSupportModule(filePath);
            } catch (err) {
                const msg = err instanceof Error ? err.message : String(err);
                log(`[installVbaTestSupport] FAILED: ${msg}`);
                recordWriteAudit({
                    command: 'xlide.installVbaTestSupport',
                    operation: 'write-module',
                    outcome: 'failed',
                    workbookPath: filePath,
                    moduleName: XLIDE_ASSERT_MODULE_NAME,
                    summary: 'Install VBA test support module: 0 changed, 1 failed',
                    error: err,
                });
                vscode.window.showErrorMessage(`XLIDE: Failed to install VBA test support: ${msg}`);
            }
        }),

        // Create a new, empty macro-enabled workbook
        registerXlideCommand('xlide.newWorkbook', async () => {
            const defaultDir = vscode.workspace.workspaceFolders?.[0]?.uri;
            const target = await vscode.window.showSaveDialog({
                title: 'XLIDE: New Macro-Enabled Workbook',
                defaultUri: defaultDir ? vscode.Uri.joinPath(defaultDir, 'NewWorkbook.xlsm') : undefined,
                filters: { 'Macro-Enabled Workbook': ['xlsm', 'xlsb'] },
            });
            if (!target) { return; }
            const filePath = target.fsPath;
            const name = path.basename(filePath);
            await vscode.window.withProgress(
                { location: vscode.ProgressLocation.Notification, title: `XLIDE: Creating "${name}"...`, cancellable: false },
                async () => {
                    try {
                        await bridge.call<{ ok: boolean; path: string }>('createWorkbook', { path: filePath });
                        log(`[newWorkbook] Created "${filePath}"`);
                        explorer.refresh();
                        void vscode.window.showInformationMessage(`XLIDE: Created "${name}".`);
                    } catch (err) {
                        const msg = err instanceof Error ? err.message : String(err);
                        log(`[newWorkbook] FAILED: ${msg}`);
                        vscode.window.showErrorMessage(`XLIDE: Failed to create workbook: ${msg}`);
                    }
                },
            );
        }),

        // Open the workbook in Excel (editable)
        registerXlideCommand('xlide.openWorkbook', async (node: XlideNode) => {
            const filePath = resolveWorkbookPath(node);
            if (!filePath) { return; }
            try {
                const attachToRunning = shouldAttachToRunningExcel();
                log(`[openWorkbook] Requested for: ${filePath}`);
                if (process.platform === 'win32') {
                    runWindowsExcel(filePath, attachToRunning, false);
                } else if (process.platform === 'darwin') {
                    cp.spawn('open', ['-a', 'Microsoft Excel', filePath]);
                } else {
                    cp.spawn('libreoffice', ['--calc', '--norestore', filePath]);
                }
            } catch (err) {
                vscode.window.showErrorMessage(`Failed to open workbook: ${err}`);
            }
        }),

        // Open the workbook in Excel (read-only)
        registerXlideCommand('xlide.openWorkbookReadOnly', async (node: XlideNode) => {
            const filePath = resolveWorkbookPath(node);
            if (!filePath) { return; }
            try {
                const attachToRunning = shouldAttachToRunningExcel();
                log(`[openWorkbookReadOnly] Requested for: ${filePath}`);
                if (process.platform === 'win32') {
                    runWindowsExcel(filePath, attachToRunning, true);
                } else if (process.platform === 'darwin') {
                    cp.spawn('open', ['-a', 'Microsoft Excel', filePath]);
                } else {
                    cp.spawn('libreoffice', ['--calc', '--norestore', '--view', filePath]);
                }
            } catch (err) {
                vscode.window.showErrorMessage(`Failed to open workbook: ${err}`);
            }
        }),

        // Detect the Sub/Function at the cursor and open the workbook, then guide to run it
        registerXlideCommand('xlide.runMacroAtCursor', async () => {
            const editor = vscode.window.activeTextEditor;
            if (!editor || !editor.document.uri.scheme.startsWith(XLIDE_SCHEME)) {
                vscode.window.showWarningMessage('XLIDE: Open a VBA module to run a macro.');
                return;
            }

            try {
                // Persist any in-editor changes first so the macro that runs
                // reflects the current source rather than the last-saved version.
                if (editor.document.isDirty) {
                    await editor.document.save();
                }

                // Decode the URI to get filePath and moduleName
                const { xlsmPath, moduleName } = decodeModuleUri(editor.document.uri);
                log(`[runMacro] Requested from module: ${moduleName} in ${xlsmPath}`);

                // Get the source code and find which Sub/Function the cursor is in
                const result = await bridge.call<{ source: string }>(
                    'readModule',
                    { path: xlsmPath, module: moduleName },
                );

                const cursorLine = editor.selection.active.line;
                const source = result.source;
                const lines = source.split('\n');

                // Find the current Sub/Function
                const procRe = /^\s*(Public|Private)?\s*(Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)/i;
                let currentProc = '';
                for (let i = cursorLine; i >= 0; i--) {
                    const match = lines[i].match(procRe);
                    if (match) {
                        currentProc = match[3];
                        break;
                    }
                }

                if (!currentProc) {
                    vscode.window.showWarningMessage('XLIDE: Cursor is not inside a Sub or Function.');
                    return;
                }

                // Open the workbook read-only
                if (process.platform === 'win32') {
                    const attachToRunning = shouldAttachToRunningExcel();
                    log(`[runMacro] attachToRunningExcel=${attachToRunning}`);
                    await runWindowsExcelMacroReadOnly(xlsmPath, `${moduleName}.${currentProc}`, attachToRunning);
                } else if (process.platform === 'darwin') {
                    cp.spawn('open', ['-a', 'Microsoft Excel', xlsmPath]);
                    vscode.window.showInformationMessage(
                        `Workbook opened. Run macro: ${moduleName}.${currentProc}`,
                    );
                } else {
                    cp.spawn('libreoffice', ['--calc', '--norestore', '--view', xlsmPath]);
                    vscode.window.showInformationMessage(
                        `Workbook opened. Run macro manually: ${moduleName}.${currentProc}`,
                    );
                }
            } catch (err) {
                showRunMacroFailure(err);
            }
        }),
    ];
}
