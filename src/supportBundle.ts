import * as path from 'path';
import type { XlideCommandLogEntry } from './xlideCommandLog';
import {
    redactSupportLogLine,
    type XlideOutputLogEntry,
} from './xlideOutputLog';
import type { XlideWriteAuditEntry } from './xlideWriteAudit';

export interface SupportBundleSetting {
    key: string;
    value: unknown;
    source: 'default' | 'machine' | 'unknown';
}

export interface SupportBundleAnalysisSummary {
    available: boolean;
    moduleType?: string;
    errorCount?: number;
    warningCount?: number;
    suppressedCount?: number;
    byCode?: Record<string, number>;
}

export interface SupportBundleWorkbookSummary {
    available: boolean;
    workbookPath?: string;
    extension?: string;
    moduleCount?: number;
    moduleTypes?: Record<string, number>;
    activeModuleType?: string;
}

export interface SupportBundleAnonymizedAnalysisModule {
    moduleIndex: number;
    moduleType: string;
    problemCount: number;
    errorCount: number;
    warningCount: number;
    byCode: Record<string, number>;
}

export interface SupportBundleAnonymizedAnalysisReport {
    included: boolean;
    unavailableReason?: 'not-requested' | 'no-active-workbook' | 'analysis-failed';
    errorCategory?: string;
    workbookExtension?: string;
    moduleCount?: number;
    problemCount?: number;
    errorCount?: number;
    warningCount?: number;
    suppressedCount?: number;
    byCode?: Record<string, number>;
    byCategory?: Record<string, number>;
    byDiagnosticKind?: Record<string, number>;
    vbeCompileEquivalentCount?: number;
    nonVbeCompileEquivalentCount?: number;
    modules?: SupportBundleAnonymizedAnalysisModule[];
}

export interface SupportBundleInput {
    generatedAt: string;
    extension: {
        id?: string;
        name?: string;
        version?: string;
    };
    vscode: {
        version: string;
        appName?: string;
    };
    runtime: {
        platform: NodeJS.Platform | string;
        arch: string;
        node: string;
    };
    workspace: {
        folderCount: number;
    };
    settings: SupportBundleSetting[];
    workbook: SupportBundleWorkbookSummary;
    analysis: SupportBundleAnalysisSummary;
    commands: XlideCommandLogEntry[];
    writeAudits?: XlideWriteAuditEntry[];
    anonymizedWorkbookAnalysisReport?: SupportBundleAnonymizedAnalysisReport;
    selectedLogs?: XlideOutputLogEntry[];
}

export interface SupportBundle {
    schemaVersion: 1;
    generatedAt: string;
    extension: SupportBundleInput['extension'];
    vscode: SupportBundleInput['vscode'];
    runtime: SupportBundleInput['runtime'];
    workspace: SupportBundleInput['workspace'];
    settings: SupportBundleSetting[];
    setup: {
        diagnosticsEnabled: boolean | undefined;
        docsEnabled: boolean | undefined;
        pythonPathConfigured: boolean;
        excelComStatus: 'available-on-windows-not-checked' | 'not-supported-on-platform';
    };
    workbook: SupportBundleWorkbookSummary;
    analysis: SupportBundleAnalysisSummary;
    anonymizedReports: {
        workbookAnalysis: SupportBundleAnonymizedAnalysisReport;
    };
    recentCommands: XlideCommandLogEntry[];
    recentWriteAudits: XlideWriteAuditEntry[];
    selectedLogs: {
        included: boolean;
        entries: XlideOutputLogEntry[];
    };
    privacy: {
        workbookSourceIncluded: false;
        pathsRedacted: true;
        commandArgumentsIncluded: false;
        writeAuditIncluded: true;
        anonymizedAnalysisReportIncluded: boolean;
        selectedLogsIncluded: boolean;
        logPathsRedacted: true;
    };
}

const PATH_SETTING_RE = /(path|folder|directory|file)$/i;
const UNAVAILABLE = 'unavailable';

export function buildSupportBundle(input: SupportBundleInput): SupportBundle {
    const settings = input.settings
        .map((setting) => ({
            ...setting,
            value: sanitizeSettingValue(setting.key, setting.value),
        }))
        .sort((a, b) => a.key.localeCompare(b.key));

    return {
        schemaVersion: 1,
        generatedAt: input.generatedAt,
        extension: input.extension,
        vscode: input.vscode,
        runtime: input.runtime,
        workspace: input.workspace,
        settings,
        setup: {
            diagnosticsEnabled: booleanSetting(settings, 'xlide.diagnostics.enabled'),
            docsEnabled: booleanSetting(settings, 'xlide.docs.enabled'),
            pythonPathConfigured: stringSettingConfigured(input.settings, 'xlide.pythonPath'),
            excelComStatus: input.runtime.platform === 'win32'
                ? 'available-on-windows-not-checked'
                : 'not-supported-on-platform',
        },
        workbook: sanitizeWorkbookSummary(input.workbook),
        analysis: input.analysis,
        anonymizedReports: {
            workbookAnalysis: input.anonymizedWorkbookAnalysisReport ?? {
                included: false,
                unavailableReason: 'not-requested',
            },
        },
        recentCommands: input.commands.slice(-25),
        recentWriteAudits: (input.writeAudits ?? []).slice(-25).map(sanitizeWriteAuditEntry),
        selectedLogs: {
            included: Boolean(input.selectedLogs),
            entries: (input.selectedLogs ?? []).slice(-50).map((entry) => ({
                timestamp: entry.timestamp,
                line: redactSupportLogLine(entry.line),
            })),
        },
        privacy: {
            workbookSourceIncluded: false,
            pathsRedacted: true,
            commandArgumentsIncluded: false,
            writeAuditIncluded: true,
            anonymizedAnalysisReportIncluded: input.anonymizedWorkbookAnalysisReport?.included === true,
            selectedLogsIncluded: Boolean(input.selectedLogs),
            logPathsRedacted: true,
        },
    };
}

export function anonymizedWorkbookAnalysisReportFromResult(result: {
    filePath: string;
    moduleCount: number;
    problems: readonly {
        moduleName: string;
        moduleType: string;
        severity: string;
        code?: string;
        category?: string;
        diagnosticKind?: string;
        vbeCompileEquivalent?: boolean;
    }[];
    errorCount: number;
    warningCount: number;
    summary: {
        byCategory: Record<string, number> | Partial<Record<string, number>>;
        byDiagnosticKind: Record<string, number> | Partial<Record<string, number>>;
        vbeCompileEquivalentCount: number;
        nonVbeCompileEquivalentCount: number;
        suppressedCount: number;
    };
}): SupportBundleAnonymizedAnalysisReport {
    const byModule = new Map<string, {
        moduleType: string;
        problemCount: number;
        errorCount: number;
        warningCount: number;
        byCode: Record<string, number>;
    }>();
    const byCode: Record<string, number> = {};
    for (const problem of result.problems) {
        byCode[problem.code ?? 'unclassified'] = (byCode[problem.code ?? 'unclassified'] ?? 0) + 1;
        const existing = byModule.get(problem.moduleName) ?? {
            moduleType: problem.moduleType,
            problemCount: 0,
            errorCount: 0,
            warningCount: 0,
            byCode: {},
        };
        existing.problemCount++;
        if (problem.severity === 'error') {
            existing.errorCount++;
        } else if (problem.severity === 'warning') {
            existing.warningCount++;
        }
        existing.byCode[problem.code ?? 'unclassified'] =
            (existing.byCode[problem.code ?? 'unclassified'] ?? 0) + 1;
        byModule.set(problem.moduleName, existing);
    }

    return {
        included: true,
        workbookExtension: path.extname(result.filePath),
        moduleCount: result.moduleCount,
        problemCount: result.problems.length,
        errorCount: result.errorCount,
        warningCount: result.warningCount,
        suppressedCount: result.summary.suppressedCount,
        byCode: sortedRecord(byCode),
        byCategory: sortedRecord(result.summary.byCategory),
        byDiagnosticKind: sortedRecord(result.summary.byDiagnosticKind),
        vbeCompileEquivalentCount: result.summary.vbeCompileEquivalentCount,
        nonVbeCompileEquivalentCount: result.summary.nonVbeCompileEquivalentCount,
        modules: [...byModule.values()]
            .map((module, index) => ({
                moduleIndex: index + 1,
                moduleType: module.moduleType,
                problemCount: module.problemCount,
                errorCount: module.errorCount,
                warningCount: module.warningCount,
                byCode: sortedRecord(module.byCode),
            }))
            .sort((a, b) => {
                if (a.moduleType !== b.moduleType) {
                    return a.moduleType.localeCompare(b.moduleType);
                }
                return a.moduleIndex - b.moduleIndex;
            }),
    };
}

export function defaultSupportBundleFileName(now = new Date()): string {
    const stamp = now.toISOString().replace(/[:.]/g, '-');
    return `xlide-support-${stamp}.json`;
}

export function supportBundleDisclosureText(bundle: SupportBundle): string {
    return [
        'XLIDE will export a redacted JSON diagnostic snapshot.',
        '',
        'Included:',
        '- Extension, VS Code, platform, Node, and workspace folder count.',
        '- XLIDE settings with path-like values redacted.',
        '- Setup states that can be determined without probing Excel.',
        '- Active workbook/module metadata and active-module analysis counts when available.',
        '- Recent XLIDE command ids, outcomes, durations, and error categories.',
        '- Recent XLIDE write-audit entries with paths redacted.',
        '',
        'Optional when explicitly selected:',
        '- Anonymized workbook analysis report with counts by rule/module type only.',
        '- Recent selected XLIDE operation log lines with paths redacted.',
        '',
        'Not included:',
        '- Workbook VBA source or cell data.',
        '- Full workbook paths or path-like setting values.',
        '- Command arguments.',
        '- Output logs unless you explicitly select the log option.',
        '',
        `Active workbook: ${workbookLine(bundle)}`,
        `Active module analysis: ${analysisLine(bundle)}`,
        `Anonymized analysis report included: ${bundle.privacy.anonymizedAnalysisReportIncluded}`,
        `Selected logs included: ${bundle.privacy.selectedLogsIncluded}`,
    ].join('\n');
}

export function supportDiagnosticsText(bundle: SupportBundle): string {
    const lines = [
        'XLIDE Diagnostics',
        `Generated: ${bundle.generatedAt}`,
        `Extension: ${displayJoin([bundle.extension.id, bundle.extension.version]) || UNAVAILABLE}`,
        `VS Code: ${displayJoin([bundle.vscode.appName, bundle.vscode.version]) || UNAVAILABLE}`,
        `Runtime: ${bundle.runtime.platform}/${bundle.runtime.arch} ${bundle.runtime.node}`,
        `Workspace folders: ${bundle.workspace.folderCount}`,
        '',
        'Setup',
        `Diagnostics enabled: ${formatDiagnosticValue(bundle.setup.diagnosticsEnabled)}`,
        `Docs enabled: ${formatDiagnosticValue(bundle.setup.docsEnabled)}`,
        `Python path configured: ${bundle.setup.pythonPathConfigured}`,
        `Excel COM status: ${bundle.setup.excelComStatus}`,
        '',
        'Active Workbook',
        workbookLine(bundle),
        '',
        'Active Module Analysis',
        analysisLine(bundle),
        '',
        'XLIDE Settings',
        ...bundle.settings.map((setting) =>
            `${setting.key} (${setting.source}): ${formatDiagnosticValue(setting.value)}`,
        ),
        '',
        'Recent Commands',
        ...recentCommandLines(bundle),
        '',
        'Recent Write Audit',
        ...writeAuditLines(bundle),
        '',
        'Anonymized Workbook Analysis Report',
        anonymizedAnalysisReportLine(bundle.anonymizedReports.workbookAnalysis),
        '',
        'Selected Logs',
        ...selectedLogLines(bundle),
        '',
        'Privacy',
        `Workbook source included: ${bundle.privacy.workbookSourceIncluded}`,
        `Paths redacted: ${bundle.privacy.pathsRedacted}`,
        `Command arguments included: ${bundle.privacy.commandArgumentsIncluded}`,
        `Write audit included: ${bundle.privacy.writeAuditIncluded}`,
        `Anonymized analysis report included: ${bundle.privacy.anonymizedAnalysisReportIncluded}`,
        `Selected logs included: ${bundle.privacy.selectedLogsIncluded}`,
        `Log paths redacted: ${bundle.privacy.logPathsRedacted}`,
    ];
    return `${lines.join('\n')}\n`;
}

export function redactPath(value: string): string {
    const ext = path.extname(value);
    return ext ? `<redacted>${ext}` : '<redacted>';
}

function sanitizeSettingValue(key: string, value: unknown): unknown {
    if (typeof value === 'string' && PATH_SETTING_RE.test(key)) {
        return value ? redactPath(value) : value;
    }
    return value;
}

function sanitizeWorkbookSummary(summary: SupportBundleWorkbookSummary): SupportBundleWorkbookSummary {
    if (!summary.available) {
        return { available: false };
    }
    return {
        ...summary,
        workbookPath: summary.workbookPath ? redactPath(summary.workbookPath) : undefined,
    };
}

function sanitizeWriteAuditEntry(entry: XlideWriteAuditEntry): XlideWriteAuditEntry {
    return {
        ...entry,
        workbookPath: entry.workbookPath ? redactPath(entry.workbookPath) : undefined,
        sourcePath: entry.sourcePath ? redactPath(entry.sourcePath) : undefined,
        targetPath: entry.targetPath ? redactPath(entry.targetPath) : undefined,
    };
}

function booleanSetting(
    settings: readonly SupportBundleSetting[],
    key: string,
): boolean | undefined {
    const value = settings.find((setting) => setting.key === key)?.value;
    return typeof value === 'boolean' ? value : undefined;
}

function stringSettingConfigured(
    settings: readonly SupportBundleSetting[],
    key: string,
): boolean {
    const value = settings.find((setting) => setting.key === key)?.value;
    return typeof value === 'string' && value.trim().length > 0;
}

function sortedRecord(
    counts: Record<string, number> | Partial<Record<string, number>>,
): Record<string, number> {
    return Object.fromEntries(
        Object.entries(counts)
            .filter((entry): entry is [string, number] => typeof entry[1] === 'number')
            .sort(([a], [b]) => a.localeCompare(b)),
    );
}

function workbookLine(bundle: SupportBundle): string {
    if (!bundle.workbook.available) {
        return UNAVAILABLE;
    }
    const moduleTypes = bundle.workbook.moduleTypes
        ? Object.entries(bundle.workbook.moduleTypes)
            .sort(([a], [b]) => a.localeCompare(b))
            .map(([name, count]) => `${name} ${count}`)
            .join(', ')
        : UNAVAILABLE;
    return [
        bundle.workbook.workbookPath ?? '<redacted>',
        `extension ${bundle.workbook.extension ?? UNAVAILABLE}`,
        `modules ${bundle.workbook.moduleCount ?? UNAVAILABLE}`,
        `moduleTypes ${moduleTypes}`,
        `activeModuleType ${bundle.workbook.activeModuleType ?? UNAVAILABLE}`,
    ].join('; ');
}

function analysisLine(bundle: SupportBundle): string {
    if (!bundle.analysis.available) {
        return UNAVAILABLE;
    }
    const byCode = bundle.analysis.byCode
        ? Object.entries(bundle.analysis.byCode)
            .sort(([a], [b]) => a.localeCompare(b))
            .map(([code, count]) => `${code} ${count}`)
            .join(', ')
        : 'none';
    return [
        `moduleType ${bundle.analysis.moduleType ?? UNAVAILABLE}`,
        `errors ${bundle.analysis.errorCount ?? 0}`,
        `warnings ${bundle.analysis.warningCount ?? 0}`,
        `suppressed ${bundle.analysis.suppressedCount ?? 0}`,
        `byCode ${byCode || 'none'}`,
    ].join('; ');
}

function recentCommandLines(bundle: SupportBundle): string[] {
    if (bundle.recentCommands.length === 0) {
        return ['none'];
    }
    return bundle.recentCommands.map((entry) => {
        const parts = [
            entry.timestamp,
            entry.command,
            entry.outcome,
        ];
        if (entry.durationMs !== undefined) {
            parts.push(`${entry.durationMs}ms`);
        }
        if (entry.errorCategory) {
            parts.push(`errorCategory=${entry.errorCategory}`);
        }
        return parts.join(' | ');
    });
}

function writeAuditLines(bundle: SupportBundle): string[] {
    if (bundle.recentWriteAudits.length === 0) {
        return ['none'];
    }
    return bundle.recentWriteAudits.map((entry) => {
        const parts = [
            entry.timestamp,
            entry.command,
            entry.operation,
            entry.outcome,
            entry.summary,
        ];
        if (entry.moduleName) {
            parts.push(`module=${entry.moduleName}`);
        }
        if (entry.errorCategory) {
            parts.push(`errorCategory=${entry.errorCategory}`);
        }
        return parts.join(' | ');
    });
}

function anonymizedAnalysisReportLine(report: SupportBundleAnonymizedAnalysisReport): string {
    if (!report.included) {
        return report.unavailableReason ?? 'not-requested';
    }
    return [
        `workbookExtension ${report.workbookExtension ?? UNAVAILABLE}`,
        `modules ${report.moduleCount ?? 0}`,
        `problems ${report.problemCount ?? 0}`,
        `errors ${report.errorCount ?? 0}`,
        `warnings ${report.warningCount ?? 0}`,
        `suppressed ${report.suppressedCount ?? 0}`,
    ].join('; ');
}

function selectedLogLines(bundle: SupportBundle): string[] {
    if (!bundle.selectedLogs.included) {
        return ['not-requested'];
    }
    if (bundle.selectedLogs.entries.length === 0) {
        return ['none'];
    }
    return bundle.selectedLogs.entries.map((entry) => `${entry.timestamp} | ${entry.line}`);
}

function displayJoin(values: readonly (string | undefined)[]): string {
    return values.filter((value): value is string => Boolean(value)).join(' ');
}

function formatDiagnosticValue(value: unknown): string {
    if (typeof value === 'string') {
        return value.length > 0 ? value : '""';
    }
    if (typeof value === 'number' || typeof value === 'boolean') {
        return String(value);
    }
    if (value === undefined) {
        return UNAVAILABLE;
    }
    if (value === null) {
        return 'null';
    }
    return JSON.stringify(value);
}
