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
