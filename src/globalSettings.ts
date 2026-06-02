import type * as vscode from 'vscode';
import {
    ANALYSIS_SEVERITIES,
    normalizeAnalysisRuleCodes,
    normalizeAnalysisVisibleSeverities,
    type AnalysisSeverityFilter,
} from './analysisSettingsCore';
import {
    DEFAULT_VBA_SMART_BLOCK_LAYOUT,
    normalizeSmartBlockLayout,
    type VbaSmartBlockLayout,
} from './vbaStructuralAnalysis';

type XlideGlobalSettingSeverity = 'error';

interface XlideGlobalSettingsProblem {
    key: string;
    message: string;
    severity: XlideGlobalSettingSeverity;
}

type XlideGlobalSettingSource = 'default' | 'machine' | 'unknown';

interface ResolvedXlideGlobalSetting<T> {
    key: `xlide.${XlideGlobalSettingKey}`;
    value: T;
    source: XlideGlobalSettingSource;
}

type XlideGlobalSettingsSnapshot = Record<string, unknown>;

const OPTION_EXPLICIT_VALUES = ['off', 'information', 'warning', 'error'] as const;
const BLOCK_LAYOUT_VALUES = ['comfy', 'compact'] as const;
const DEFAULT_DOC_METADATA_GLOB = '**/*.vbref.xml';
const XLIDE_GLOBAL_SETTING_KEYS = [
    'analysis.untrackedRules',
    'analysis.visibleSeverities',
    'attachToRunningExcel',
    'diagnostics.enabled',
    'diagnostics.optionExplicit',
    'docs.enabled',
    'docs.metadataGlob',
    'editor.blockLayout',
    'pythonPath',
] as const;
type XlideGlobalSettingKey = typeof XLIDE_GLOBAL_SETTING_KEYS[number];
type XlideOptionExplicitSetting = typeof OPTION_EXPLICIT_VALUES[number];

function normalizeXlideOptionExplicitSetting(value: unknown): XlideOptionExplicitSetting {
    return isXlideOptionExplicitSetting(value)
        ? value
        : 'warning';
}

function xlideAnalysisVisibleSeveritiesFromConfig(
    config: vscode.WorkspaceConfiguration,
): ResolvedXlideGlobalSetting<AnalysisSeverityFilter[]> {
    return resolveXlideGlobalSetting(
        config,
        'analysis.visibleSeverities',
        [...ANALYSIS_SEVERITIES],
        normalizeAnalysisVisibleSeverities,
    );
}

function xlideAnalysisUntrackedRulesFromConfig(
    config: vscode.WorkspaceConfiguration,
): ResolvedXlideGlobalSetting<string[]> {
    return resolveXlideGlobalSetting(config, 'analysis.untrackedRules', [], normalizeAnalysisRuleCodes);
}

function xlideAttachToRunningExcelFromConfig(
    config: vscode.WorkspaceConfiguration,
): ResolvedXlideGlobalSetting<boolean> {
    return resolveXlideGlobalSetting(config, 'attachToRunningExcel', true, normalizeBoolean(true));
}

function xlideDiagnosticsEnabledFromConfig(
    config: vscode.WorkspaceConfiguration,
): ResolvedXlideGlobalSetting<boolean> {
    return resolveXlideGlobalSetting(config, 'diagnostics.enabled', true, normalizeBoolean(true));
}

function xlideOptionExplicitFromConfig(
    config: vscode.WorkspaceConfiguration,
): ResolvedXlideGlobalSetting<XlideOptionExplicitSetting> {
    return resolveXlideGlobalSetting(
        config,
        'diagnostics.optionExplicit',
        'warning',
        normalizeXlideOptionExplicitSetting,
    );
}

function xlideDocsEnabledFromConfig(
    config: vscode.WorkspaceConfiguration,
): ResolvedXlideGlobalSetting<boolean> {
    return resolveXlideGlobalSetting(config, 'docs.enabled', true, normalizeBoolean(true));
}

function xlideDocsMetadataGlobFromConfig(
    config: vscode.WorkspaceConfiguration,
): ResolvedXlideGlobalSetting<string> {
    return resolveXlideGlobalSetting(
        config,
        'docs.metadataGlob',
        DEFAULT_DOC_METADATA_GLOB,
        (value) => normalizeNonEmptyString(value, DEFAULT_DOC_METADATA_GLOB),
    );
}

function xlideEditorBlockLayoutFromConfig(
    config: vscode.WorkspaceConfiguration,
): ResolvedXlideGlobalSetting<VbaSmartBlockLayout> {
    return resolveXlideGlobalSetting(
        config,
        'editor.blockLayout',
        DEFAULT_VBA_SMART_BLOCK_LAYOUT,
        normalizeSmartBlockLayout,
    );
}

function xlidePythonPathFromConfig(
    config: vscode.WorkspaceConfiguration,
): ResolvedXlideGlobalSetting<string> {
    return resolveXlideGlobalSetting(config, 'pythonPath', '', (value) =>
        typeof value === 'string' ? value.trim() : '',
    );
}

function resolvedXlideGlobalSettingsFromConfig(
    config: vscode.WorkspaceConfiguration,
): ResolvedXlideGlobalSetting<unknown>[] {
    return [
        xlideAnalysisUntrackedRulesFromConfig(config),
        xlideAnalysisVisibleSeveritiesFromConfig(config),
        xlideAttachToRunningExcelFromConfig(config),
        xlideDiagnosticsEnabledFromConfig(config),
        xlideOptionExplicitFromConfig(config),
        xlideDocsEnabledFromConfig(config),
        xlideDocsMetadataGlobFromConfig(config),
        xlideEditorBlockLayoutFromConfig(config),
        xlidePythonPathFromConfig(config),
    ];
}

function isXlideOptionExplicitSetting(value: unknown): value is XlideOptionExplicitSetting {
    return typeof value === 'string' &&
        (OPTION_EXPLICIT_VALUES as readonly string[]).includes(value);
}

function validateXlideGlobalSettingsValues(values: XlideGlobalSettingsSnapshot): XlideGlobalSettingsProblem[] {
    const problems: XlideGlobalSettingsProblem[] = [];

    expectString(values, problems, 'pythonPath');
    expectBoolean(values, problems, 'attachToRunningExcel');
    expectBoolean(values, problems, 'diagnostics.enabled');
    expectEnum(values, problems, 'diagnostics.optionExplicit', OPTION_EXPLICIT_VALUES);
    expectStringArrayEnum(values, problems, 'analysis.visibleSeverities', ANALYSIS_SEVERITIES);
    expectNonEmptyStringArray(values, problems, 'analysis.untrackedRules');
    expectEnum(values, problems, 'editor.blockLayout', BLOCK_LAYOUT_VALUES);
    expectBoolean(values, problems, 'docs.enabled');
    expectString(values, problems, 'docs.metadataGlob');

    return problems;
}

function validateXlideGlobalSettingsFromConfig(
    config: vscode.WorkspaceConfiguration,
): XlideGlobalSettingsProblem[] {
    return validateXlideGlobalSettingsValues({
        pythonPath: config.get<unknown>('pythonPath'),
        attachToRunningExcel: config.get<unknown>('attachToRunningExcel'),
        'diagnostics.enabled': config.get<unknown>('diagnostics.enabled'),
        'diagnostics.optionExplicit': config.get<unknown>('diagnostics.optionExplicit'),
        'analysis.visibleSeverities': config.get<unknown>('analysis.visibleSeverities'),
        'analysis.untrackedRules': config.get<unknown>('analysis.untrackedRules'),
        'editor.blockLayout': config.get<unknown>('editor.blockLayout'),
        'docs.enabled': config.get<unknown>('docs.enabled'),
        'docs.metadataGlob': config.get<unknown>('docs.metadataGlob'),
    });
}

function resolveXlideGlobalSetting<T>(
    config: vscode.WorkspaceConfiguration,
    key: XlideGlobalSettingKey,
    defaultValue: T,
    normalize: (value: unknown) => T,
): ResolvedXlideGlobalSetting<T> {
    return {
        key: `xlide.${key}`,
        value: normalize(config.get<unknown>(key, defaultValue)),
        source: xlideGlobalSettingSource(
            typeof config.inspect === 'function' ? config.inspect(key) : undefined,
        ),
    };
}

function xlideGlobalSettingSource(
    inspect: { globalValue?: unknown } | undefined,
): XlideGlobalSettingSource {
    if (!inspect) {
        return 'unknown';
    }
    return inspect.globalValue === undefined ? 'default' : 'machine';
}

function normalizeBoolean(fallback: boolean): (value: unknown) => boolean {
    return (value) => typeof value === 'boolean' ? value : fallback;
}

function normalizeNonEmptyString(value: unknown, fallback: string): string {
    return typeof value === 'string' && value.trim().length > 0 ? value.trim() : fallback;
}

function expectString(
    values: XlideGlobalSettingsSnapshot,
    problems: XlideGlobalSettingsProblem[],
    key: string,
): void {
    if (typeof values[key] !== 'string') {
        problems.push(problem(key, `Expected "xlide.${key}" to be a string.`));
    }
}

function expectBoolean(
    values: XlideGlobalSettingsSnapshot,
    problems: XlideGlobalSettingsProblem[],
    key: string,
): void {
    if (typeof values[key] !== 'boolean') {
        problems.push(problem(key, `Expected "xlide.${key}" to be true or false.`));
    }
}

function expectEnum<T extends readonly string[]>(
    values: XlideGlobalSettingsSnapshot,
    problems: XlideGlobalSettingsProblem[],
    key: string,
    allowed: T,
): void {
    const value = values[key];
    if (typeof value === 'string' && allowed.includes(value)) {
        return;
    }
    problems.push(problem(key, `Expected "xlide.${key}" to be one of: ${allowed.join(', ')}.`));
}

function expectStringArrayEnum<T extends readonly string[]>(
    values: XlideGlobalSettingsSnapshot,
    problems: XlideGlobalSettingsProblem[],
    key: string,
    allowed: T,
): void {
    const value = values[key];
    if (!Array.isArray(value)) {
        problems.push(problem(key, `Expected "xlide.${key}" to be an array.`));
        return;
    }
    const invalid = value.find((entry) => typeof entry !== 'string' || !allowed.includes(entry));
    if (invalid !== undefined) {
        problems.push(problem(key, `Expected "xlide.${key}" entries to be one of: ${allowed.join(', ')}.`));
    }
}

function expectNonEmptyStringArray(
    values: XlideGlobalSettingsSnapshot,
    problems: XlideGlobalSettingsProblem[],
    key: string,
): void {
    const value = values[key];
    if (!Array.isArray(value)) {
        problems.push(problem(key, `Expected "xlide.${key}" to be an array.`));
        return;
    }
    if (value.some((entry) => typeof entry !== 'string' || entry.trim().length === 0)) {
        problems.push(problem(key, `Expected "xlide.${key}" entries to be non-empty strings.`));
    }
}

function problem(key: string, message: string): XlideGlobalSettingsProblem {
    return {
        key: `xlide.${key}`,
        message,
        severity: 'error',
    };
}

export {
    DEFAULT_DOC_METADATA_GLOB,
    XLIDE_GLOBAL_SETTING_KEYS,
    type ResolvedXlideGlobalSetting,
    type XlideGlobalSettingKey,
    type XlideGlobalSettingSource,
    type XlideGlobalSettingsProblem,
    normalizeXlideOptionExplicitSetting,
    resolvedXlideGlobalSettingsFromConfig,
    validateXlideGlobalSettingsFromConfig,
    validateXlideGlobalSettingsValues,
    xlideAnalysisUntrackedRulesFromConfig,
    xlideAnalysisVisibleSeveritiesFromConfig,
    xlideAttachToRunningExcelFromConfig,
    xlideDiagnosticsEnabledFromConfig,
    xlideDocsEnabledFromConfig,
    xlideDocsMetadataGlobFromConfig,
    xlideEditorBlockLayoutFromConfig,
    xlideOptionExplicitFromConfig,
    xlidePythonPathFromConfig,
};
