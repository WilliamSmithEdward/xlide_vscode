import type * as vscode from 'vscode';
import { ANALYSIS_SEVERITIES } from './analysisSettingsCore';

type XlideGlobalSettingSeverity = 'error';

interface XlideGlobalSettingsProblem {
    key: string;
    message: string;
    severity: XlideGlobalSettingSeverity;
}

type XlideGlobalSettingsSnapshot = Record<string, unknown>;

const OPTION_EXPLICIT_VALUES = ['off', 'hint', 'information', 'warning', 'error'] as const;
const BLOCK_LAYOUT_VALUES = ['comfy', 'compact'] as const;
type XlideOptionExplicitSetting = typeof OPTION_EXPLICIT_VALUES[number];

function normalizeXlideOptionExplicitSetting(value: unknown): XlideOptionExplicitSetting {
    return isXlideOptionExplicitSetting(value)
        ? value
        : 'warning';
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
    type XlideGlobalSettingsProblem,
    normalizeXlideOptionExplicitSetting,
    validateXlideGlobalSettingsFromConfig,
    validateXlideGlobalSettingsValues,
};
