import type * as vscode from 'vscode';
import {
    ANALYSIS_SEVERITIES,
    allowedAnalysisRuleSeverityOverrides,
    normalizeAnalysisRuleCodes,
    normalizeAnalysisRuleCode,
    normalizeAnalysisRuleSeverityOverrides,
    normalizeAnalysisVisibleSeverities,
    type AnalysisRuleSeverityOverride,
    type AnalysisRuleSeverityOverrides,
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

const BLOCK_LAYOUT_VALUES = ['comfy', 'compact'] as const;
const DEFAULT_DOC_METADATA_GLOB = '**/*.vbref.xml';
const XLIDE_GLOBAL_SETTING_KEYS = [
    'analysis.ruleSeverityOverrides',
    'analysis.untrackedRules',
    'analysis.visibleSeverities',
    'attachToRunningExcel',
    'diagnostics.enabled',
    'docs.enabled',
    'docs.metadataGlob',
    'editor.blockLayout',
    'pythonPath',
] as const;
type XlideGlobalSettingKey = typeof XLIDE_GLOBAL_SETTING_KEYS[number];

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

function xlideAnalysisRuleSeveritiesFromConfig(
    config: vscode.WorkspaceConfiguration,
): ResolvedXlideGlobalSetting<AnalysisRuleSeverityOverrides> {
    return resolveXlideGlobalSetting(
        config,
        'analysis.ruleSeverityOverrides',
        {},
        normalizeAnalysisRuleSeverityOverrides,
    );
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
        xlideAnalysisRuleSeveritiesFromConfig(config),
        xlideAnalysisUntrackedRulesFromConfig(config),
        xlideAnalysisVisibleSeveritiesFromConfig(config),
        xlideAttachToRunningExcelFromConfig(config),
        xlideDiagnosticsEnabledFromConfig(config),
        xlideDocsEnabledFromConfig(config),
        xlideDocsMetadataGlobFromConfig(config),
        xlideEditorBlockLayoutFromConfig(config),
        xlidePythonPathFromConfig(config),
    ];
}

function validateXlideGlobalSettingsValues(values: XlideGlobalSettingsSnapshot): XlideGlobalSettingsProblem[] {
    const problems: XlideGlobalSettingsProblem[] = [];

    expectString(values, problems, 'pythonPath');
    expectBoolean(values, problems, 'attachToRunningExcel');
    expectBoolean(values, problems, 'diagnostics.enabled');
    expectRuleSeverityOverrides(values, problems, 'analysis.ruleSeverityOverrides');
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
        'analysis.ruleSeverityOverrides': config.get<unknown>('analysis.ruleSeverityOverrides'),
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

function expectRuleSeverityOverrides(
    values: XlideGlobalSettingsSnapshot,
    problems: XlideGlobalSettingsProblem[],
    key: string,
): void {
    const value = values[key];
    if (!value || typeof value !== 'object' || Array.isArray(value)) {
        problems.push(problem(key, `Expected "xlide.${key}" to be an object keyed by analysis rule code.`));
        return;
    }
    for (const [rawCode, rawSeverity] of Object.entries(value)) {
        const code = normalizeAnalysisRuleCode(rawCode);
        const allowed = allowedAnalysisRuleSeverityOverrides(code);
        if (!code || allowed.length === 0) {
            problems.push(problem(
                key,
                `Expected "xlide.${key}.${rawCode}" to target a known analysis rule that permits severity overrides.`,
            ));
            continue;
        }
        if (typeof rawSeverity !== 'string' || !allowed.includes(rawSeverity as AnalysisRuleSeverityOverride)) {
            problems.push(problem(
                key,
                `Expected "xlide.${key}.${rawCode}" to be one of: ${allowed.join(', ')}.`,
            ));
        }
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
    resolvedXlideGlobalSettingsFromConfig,
    validateXlideGlobalSettingsFromConfig,
    validateXlideGlobalSettingsValues,
    xlideAnalysisRuleSeveritiesFromConfig,
    xlideAnalysisUntrackedRulesFromConfig,
    xlideAnalysisVisibleSeveritiesFromConfig,
    xlideAttachToRunningExcelFromConfig,
    xlideDiagnosticsEnabledFromConfig,
    xlideDocsEnabledFromConfig,
    xlideDocsMetadataGlobFromConfig,
    xlideEditorBlockLayoutFromConfig,
    xlidePythonPathFromConfig,
};
