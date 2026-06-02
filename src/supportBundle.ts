import * as path from 'path';
import type { XlideCommandLogEntry } from './xlideCommandLog';

export interface SupportBundleSetting {
    key: string;
    value: unknown;
    source: 'default' | 'global' | 'workspace' | 'workspaceFolder' | 'unknown';
}

export interface SupportBundleLintSummary {
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
    lint: SupportBundleLintSummary;
    commands: XlideCommandLogEntry[];
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
    lint: SupportBundleLintSummary;
    recentCommands: XlideCommandLogEntry[];
    privacy: {
        workbookSourceIncluded: false;
        pathsRedacted: true;
        commandArgumentsIncluded: false;
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
        lint: input.lint,
        recentCommands: input.commands.slice(-25),
        privacy: {
            workbookSourceIncluded: false,
            pathsRedacted: true,
            commandArgumentsIncluded: false,
        },
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
        '- Active workbook/module metadata and active-module lint counts when available.',
        '- Recent XLIDE command ids, outcomes, durations, and error categories.',
        '',
        'Not included:',
        '- Workbook VBA source or cell data.',
        '- Full workbook paths or path-like setting values.',
        '- Command arguments.',
        '- Output logs unless a future opt-in export explicitly adds them.',
        '',
        `Active workbook: ${workbookLine(bundle)}`,
        `Active module lint: ${lintLine(bundle)}`,
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
        'Active Module Lint',
        lintLine(bundle),
        '',
        'XLIDE Settings',
        ...bundle.settings.map((setting) =>
            `${setting.key} (${setting.source}): ${formatDiagnosticValue(setting.value)}`,
        ),
        '',
        'Recent Commands',
        ...recentCommandLines(bundle),
        '',
        'Privacy',
        `Workbook source included: ${bundle.privacy.workbookSourceIncluded}`,
        `Paths redacted: ${bundle.privacy.pathsRedacted}`,
        `Command arguments included: ${bundle.privacy.commandArgumentsIncluded}`,
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

function lintLine(bundle: SupportBundle): string {
    if (!bundle.lint.available) {
        return UNAVAILABLE;
    }
    const byCode = bundle.lint.byCode
        ? Object.entries(bundle.lint.byCode)
            .sort(([a], [b]) => a.localeCompare(b))
            .map(([code, count]) => `${code} ${count}`)
            .join(', ')
        : 'none';
    return [
        `moduleType ${bundle.lint.moduleType ?? UNAVAILABLE}`,
        `errors ${bundle.lint.errorCount ?? 0}`,
        `warnings ${bundle.lint.warningCount ?? 0}`,
        `suppressed ${bundle.lint.suppressedCount ?? 0}`,
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
