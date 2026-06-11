import type * as vscode from 'vscode';
import {
    ANALYSIS_SEVERITIES,
    normalizeKnownAnalysisRuleCodes,
    normalizeAnalysisRuleCodes,
    normalizeAnalysisRuleCode,
    normalizeAnalysisRuleSeverityOverrides,
    normalizeAnalysisVisibleSeverities,
    planAnalysisRuleTrackingUpdate,
    validateAnalysisRuleSeverityOverrideEntries,
    type AnalysisRuleTrackingUpdate,
    type AnalysisRuleSeverityOverrides,
    type AnalysisSeverityFilter,
} from './analysisSettingsCore';
import {
    DEFAULT_VBA_SMART_BLOCK_LAYOUT,
    normalizeSmartBlockLayout,
    type VbaSmartBlockLayout,
} from './vbaSmartEnter';

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

interface XlideGlobalSettingUpdateResult<T = unknown> {
    key: `xlide.${XlideGlobalSettingKey}`;
    value: T;
    changed: boolean;
}

type XlideGlobalSettingsSnapshot = Record<string, unknown>;

interface XlideGlobalSettingValues {
    'pythonPath': string;
    'attachToRunningExcel': boolean;
    'diagnostics.enabled': boolean;
    'editor.blockLayout': VbaSmartBlockLayout;
    'docs.enabled': boolean;
    'docs.metadataGlob': string;
    'analysis.visibleSeverities': AnalysisSeverityFilter[];
    'analysis.untrackedRules': string[];
    'analysis.ruleSeverityOverrides': AnalysisRuleSeverityOverrides;
    'performance.trace': boolean;
}

type XlideGlobalSettingKey = keyof XlideGlobalSettingValues;
type XlideGlobalSettingSection = 'runtime' | 'editor' | 'docs' | 'analysis';

type XlideGlobalSettingControl =
    | { kind: 'text' }
    | { kind: 'boolean' }
    | { kind: 'enum'; values: readonly string[] }
    | { kind: 'severityFilter' }
    | { kind: 'rulePicker' }
    | { kind: 'ruleSeverityOverrides' };

interface XlideGlobalSettingWebviewCard {
    section: XlideGlobalSettingSection;
    label: string;
    control: XlideGlobalSettingControl;
}

interface XlideGlobalSettingCard extends XlideGlobalSettingWebviewCard {
    key: XlideGlobalSettingKey;
}

interface XlideGlobalSettingSchema<T> {
    defaultValue: () => T;
    normalize: (value: unknown) => T;
    validate: (
        values: XlideGlobalSettingsSnapshot,
        problems: XlideGlobalSettingsProblem[],
        key: XlideGlobalSettingKey,
    ) => void;
    /** contributes.configuration fragment minus scope, default, and description. */
    manifest: Record<string, unknown>;
    webviewCard?: XlideGlobalSettingWebviewCard;
}

const BLOCK_LAYOUT_VALUES = ['comfy', 'compact'] as const;
const RULE_SEVERITY_OVERRIDE_VALUES = ['off', 'warning'] as const;
const DEFAULT_DOC_METADATA_GLOB = '**/*.vbref.xml';

// Single source of truth for every xlide.* global setting, declared in
// settings-page display order. The package.json contributes.configuration
// block is asserted against xlideGlobalSettingManifest() by
// tests/globalSettingsManifest.test.ts.
const XLIDE_GLOBAL_SETTINGS: {
    [K in XlideGlobalSettingKey]: XlideGlobalSettingSchema<XlideGlobalSettingValues[K]>;
} = {
    'pythonPath': {
        defaultValue: () => '',
        normalize: (value) => typeof value === 'string' ? value.trim() : '',
        validate: expectString,
        manifest: { type: 'string' },
        webviewCard: { section: 'runtime', label: 'Python Path', control: { kind: 'text' } },
    },
    'attachToRunningExcel': {
        defaultValue: () => true,
        normalize: normalizeBoolean(true),
        validate: expectBoolean,
        manifest: { type: 'boolean' },
        webviewCard: { section: 'runtime', label: 'Attach To Running Excel', control: { kind: 'boolean' } },
    },
    'diagnostics.enabled': {
        defaultValue: () => true,
        normalize: normalizeBoolean(true),
        validate: expectBoolean,
        manifest: { type: 'boolean' },
        webviewCard: { section: 'editor', label: 'Diagnostics Enabled', control: { kind: 'boolean' } },
    },
    'editor.blockLayout': {
        defaultValue: () => DEFAULT_VBA_SMART_BLOCK_LAYOUT,
        normalize: normalizeSmartBlockLayout,
        validate: (values, problems, key) => expectEnum(values, problems, key, BLOCK_LAYOUT_VALUES),
        manifest: { type: 'string', enum: BLOCK_LAYOUT_VALUES },
        webviewCard: {
            section: 'editor',
            label: 'Editor Block Layout',
            control: { kind: 'enum', values: BLOCK_LAYOUT_VALUES },
        },
    },
    'docs.enabled': {
        defaultValue: () => true,
        normalize: normalizeBoolean(true),
        validate: expectBoolean,
        manifest: { type: 'boolean' },
        webviewCard: { section: 'docs', label: 'Docs Enabled', control: { kind: 'boolean' } },
    },
    'docs.metadataGlob': {
        defaultValue: () => DEFAULT_DOC_METADATA_GLOB,
        normalize: (value) => normalizeNonEmptyString(value, DEFAULT_DOC_METADATA_GLOB),
        validate: expectString,
        manifest: { type: 'string' },
        webviewCard: { section: 'docs', label: 'Docs Metadata Glob', control: { kind: 'text' } },
    },
    'analysis.visibleSeverities': {
        defaultValue: () => [...ANALYSIS_SEVERITIES],
        normalize: normalizeAnalysisVisibleSeverities,
        validate: (values, problems, key) => expectStringArrayEnum(values, problems, key, ANALYSIS_SEVERITIES),
        manifest: { type: 'array', items: { type: 'string', enum: ANALYSIS_SEVERITIES } },
        webviewCard: { section: 'analysis', label: 'Visible Severities', control: { kind: 'severityFilter' } },
    },
    'analysis.untrackedRules': {
        defaultValue: () => [],
        normalize: normalizeKnownAnalysisRuleCodes,
        validate: expectKnownAnalysisRuleCodeArray,
        manifest: { type: 'array', items: { type: 'string' } },
        webviewCard: { section: 'analysis', label: 'Globally Untracked Rules', control: { kind: 'rulePicker' } },
    },
    'analysis.ruleSeverityOverrides': {
        defaultValue: () => ({}),
        normalize: normalizeAnalysisRuleSeverityOverrides,
        validate: expectRuleSeverityOverrides,
        manifest: { type: 'object', additionalProperties: { type: 'string', enum: RULE_SEVERITY_OVERRIDE_VALUES } },
        webviewCard: {
            section: 'analysis',
            label: 'Rule Severity Overrides',
            control: { kind: 'ruleSeverityOverrides' },
        },
    },
    'performance.trace': {
        defaultValue: () => false,
        normalize: normalizeBoolean(false),
        validate: expectBoolean,
        manifest: { type: 'boolean' },
    },
};

const XLIDE_GLOBAL_SETTING_KEYS = xlideGlobalSettingRegistryKeys().sort();

function xlideGlobalSettingRegistryKeys(): XlideGlobalSettingKey[] {
    return Object.keys(XLIDE_GLOBAL_SETTINGS) as XlideGlobalSettingKey[];
}

function xlideGlobalSettingCards(): XlideGlobalSettingCard[] {
    return xlideGlobalSettingRegistryKeys().flatMap((key) => {
        const card = XLIDE_GLOBAL_SETTINGS[key].webviewCard;
        return card ? [{ key, ...card }] : [];
    });
}

function xlideGlobalSettingManifest(): Record<string, Record<string, unknown>> {
    const manifest: Record<string, Record<string, unknown>> = {};
    for (const key of XLIDE_GLOBAL_SETTING_KEYS) {
        const schema = XLIDE_GLOBAL_SETTINGS[key];
        manifest[`xlide.${key}`] = {
            ...schema.manifest,
            scope: 'machine',
            default: schema.defaultValue(),
        };
    }
    return manifest;
}

function xlideGlobalSettingFromConfig<K extends XlideGlobalSettingKey>(
    config: vscode.WorkspaceConfiguration,
    key: K,
): ResolvedXlideGlobalSetting<XlideGlobalSettingValues[K]> {
    const schema = XLIDE_GLOBAL_SETTINGS[key];
    return {
        key: `xlide.${key}`,
        value: schema.normalize(config.get<unknown>(key, schema.defaultValue())),
        source: xlideGlobalSettingSource(
            typeof config.inspect === 'function' ? config.inspect(key) : undefined,
        ),
    };
}

function xlideAnalysisVisibleSeveritiesFromConfig(config: vscode.WorkspaceConfiguration) {
    return xlideGlobalSettingFromConfig(config, 'analysis.visibleSeverities');
}

function xlideAnalysisUntrackedRulesFromConfig(config: vscode.WorkspaceConfiguration) {
    return xlideGlobalSettingFromConfig(config, 'analysis.untrackedRules');
}

function xlideAnalysisRuleSeveritiesFromConfig(config: vscode.WorkspaceConfiguration) {
    return xlideGlobalSettingFromConfig(config, 'analysis.ruleSeverityOverrides');
}

function xlideAttachToRunningExcelFromConfig(config: vscode.WorkspaceConfiguration) {
    return xlideGlobalSettingFromConfig(config, 'attachToRunningExcel');
}

function xlideDiagnosticsEnabledFromConfig(config: vscode.WorkspaceConfiguration) {
    return xlideGlobalSettingFromConfig(config, 'diagnostics.enabled');
}

function xlideDocsEnabledFromConfig(config: vscode.WorkspaceConfiguration) {
    return xlideGlobalSettingFromConfig(config, 'docs.enabled');
}

function xlideDocsMetadataGlobFromConfig(config: vscode.WorkspaceConfiguration) {
    return xlideGlobalSettingFromConfig(config, 'docs.metadataGlob');
}

function xlideEditorBlockLayoutFromConfig(config: vscode.WorkspaceConfiguration) {
    return xlideGlobalSettingFromConfig(config, 'editor.blockLayout');
}

function xlidePythonPathFromConfig(config: vscode.WorkspaceConfiguration) {
    return xlideGlobalSettingFromConfig(config, 'pythonPath');
}

function xlidePerformanceTraceFromConfig(config: vscode.WorkspaceConfiguration) {
    return xlideGlobalSettingFromConfig(config, 'performance.trace');
}

function resolvedXlideGlobalSettingsFromConfig(
    config: vscode.WorkspaceConfiguration,
): ResolvedXlideGlobalSetting<unknown>[] {
    return XLIDE_GLOBAL_SETTING_KEYS.map((key) => xlideGlobalSettingFromConfig(config, key));
}

function validateXlideGlobalSettingsValues(values: XlideGlobalSettingsSnapshot): XlideGlobalSettingsProblem[] {
    const problems: XlideGlobalSettingsProblem[] = [];
    for (const key of XLIDE_GLOBAL_SETTING_KEYS) {
        XLIDE_GLOBAL_SETTINGS[key].validate(values, problems, key);
    }
    return problems;
}

function validateXlideGlobalSettingsFromConfig(
    config: vscode.WorkspaceConfiguration,
): XlideGlobalSettingsProblem[] {
    const snapshot: XlideGlobalSettingsSnapshot = {};
    for (const key of XLIDE_GLOBAL_SETTING_KEYS) {
        snapshot[key] = config.get<unknown>(key);
    }
    return validateXlideGlobalSettingsValues(snapshot);
}

async function setXlideGlobalAnalysisRuleTracked(
    config: vscode.WorkspaceConfiguration,
    code: string | undefined,
    tracked: boolean,
): Promise<AnalysisRuleTrackingUpdate> {
    const normalized = normalizeAnalysisRuleCode(code);
    const current = normalizeKnownAnalysisRuleCodes(xlideAnalysisUntrackedRulesFromConfig(config).value);
    if (!normalized || normalizeKnownAnalysisRuleCodes([normalized]).length === 0) {
        return {
            tracked,
            changed: false,
            untrackedRules: current,
        };
    }

    const update = planAnalysisRuleTrackingUpdate(current, normalized, tracked);
    if (update.changed) {
        await config.update('analysis.untrackedRules', update.untrackedRules, true);
    }
    return update;
}

async function setXlideGlobalAnalysisRuleSeverityOverride(
    config: vscode.WorkspaceConfiguration,
    code: string | undefined,
    severity: unknown,
): Promise<XlideGlobalSettingUpdateResult<AnalysisRuleSeverityOverrides>> {
    const normalized = normalizeAnalysisRuleCode(code);
    const current = xlideAnalysisRuleSeveritiesFromConfig(config).value;
    if (!normalized) {
        return {
            key: 'xlide.analysis.ruleSeverityOverrides',
            value: current,
            changed: false,
        };
    }
    const next = normalizeAnalysisRuleSeverityOverrides({
        ...current,
        [normalized]: severity,
    });
    return updateXlideGlobalSetting(config, 'analysis.ruleSeverityOverrides', next);
}

async function clearXlideGlobalAnalysisRuleSeverityOverride(
    config: vscode.WorkspaceConfiguration,
    code: string | undefined,
): Promise<XlideGlobalSettingUpdateResult<AnalysisRuleSeverityOverrides>> {
    const normalized = normalizeAnalysisRuleCode(code);
    const current = xlideAnalysisRuleSeveritiesFromConfig(config).value;
    if (!normalized || !(normalized in current)) {
        return {
            key: 'xlide.analysis.ruleSeverityOverrides',
            value: current,
            changed: false,
        };
    }
    const next = { ...current };
    delete next[normalized];
    return updateXlideGlobalSetting(
        config,
        'analysis.ruleSeverityOverrides',
        normalizeAnalysisRuleSeverityOverrides(next),
    );
}

async function setXlideGlobalSettingValue(
    config: vscode.WorkspaceConfiguration,
    key: XlideGlobalSettingKey,
    value: unknown,
): Promise<XlideGlobalSettingUpdateResult> {
    return updateXlideGlobalSetting(config, key, normalizeXlideGlobalSettingValue(key, value));
}

async function resetXlideGlobalSettingValue(
    config: vscode.WorkspaceConfiguration,
    key: XlideGlobalSettingKey,
): Promise<XlideGlobalSettingUpdateResult> {
    const inspect = typeof config.inspect === 'function' ? config.inspect(key) : undefined;
    const changed = inspect?.globalValue !== undefined;
    if (changed) {
        await config.update(key, undefined, true);
    }
    return {
        key: `xlide.${key}`,
        value: normalizeXlideGlobalSettingValue(key, undefined),
        changed,
    };
}

async function updateXlideGlobalSetting<T>(
    config: vscode.WorkspaceConfiguration,
    key: XlideGlobalSettingKey,
    value: T,
): Promise<XlideGlobalSettingUpdateResult<T>> {
    const current = xlideGlobalSettingFromConfig(config, key).value;
    const changed = !equivalentSettingValue(current, value);
    if (changed) {
        await config.update(key, value, true);
    }
    return {
        key: `xlide.${key}`,
        value,
        changed,
    };
}

function normalizeXlideGlobalSettingValue<K extends XlideGlobalSettingKey>(
    key: K,
    value: unknown,
): XlideGlobalSettingValues[K] {
    return XLIDE_GLOBAL_SETTINGS[key].normalize(value);
}

function equivalentSettingValue(left: unknown, right: unknown): boolean {
    return JSON.stringify(left) === JSON.stringify(right);
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

function expectKnownAnalysisRuleCodeArray(
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
        return;
    }
    const invalid = value.find((entry) => normalizeKnownAnalysisRuleCodes([entry]).length === 0);
    if (invalid !== undefined) {
        problems.push(problem(key, `Expected "xlide.${key}" entries to be known analysis rule codes.`));
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
    validateAnalysisRuleSeverityOverrideEntries(value as Record<string, unknown>, (rawCode, requirement) => {
        problems.push(problem(key, `Expected "xlide.${key}.${rawCode}" ${requirement}`));
    });
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
    type XlideGlobalSettingCard,
    type XlideGlobalSettingControl,
    type XlideGlobalSettingKey,
    type XlideGlobalSettingSection,
    type XlideGlobalSettingSource,
    type XlideGlobalSettingUpdateResult,
    type XlideGlobalSettingsProblem,
    clearXlideGlobalAnalysisRuleSeverityOverride,
    resolvedXlideGlobalSettingsFromConfig,
    resetXlideGlobalSettingValue,
    setXlideGlobalAnalysisRuleSeverityOverride,
    setXlideGlobalAnalysisRuleTracked,
    setXlideGlobalSettingValue,
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
    xlideGlobalSettingCards,
    xlideGlobalSettingManifest,
    xlidePerformanceTraceFromConfig,
    xlidePythonPathFromConfig,
};
