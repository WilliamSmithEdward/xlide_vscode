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
    'excelIntegration.coordinationMode': ExcelCoordinationMode;
    'excelIntegration.trackOpenedWorkbooks': boolean;
    'excelIntegration.reopenAfterClose': boolean;
    'excelIntegration.reopenMode': ExcelReopenMode;
    'excelIntegration.reopenReadOnlyAfterSave': boolean;
    'diagnostics.enabled': boolean;
    'editor.blockLayout': VbaSmartBlockLayout;
    'editor.continueCommentOnNewline': boolean;
    'editor.mirrorCommentSpacing': boolean;
    'docs.enabled': boolean;
    'docs.metadataGlob': string;
    'analysis.visibleSeverities': AnalysisSeverityFilter[];
    'analysis.untrackedRules': string[];
    'analysis.ruleSeverityOverrides': AnalysisRuleSeverityOverrides;
    'performance.trace': boolean;
}

type XlideGlobalSettingKey = keyof XlideGlobalSettingValues;
type XlideGlobalSettingSection = 'runtime' | 'excel' | 'editor' | 'docs' | 'analysis';

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
    /** Plain-language hover help shown in the settings panel info bubble. */
    description?: string;
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

const EXCEL_COORDINATION_MODE_VALUES = ['block', 'closeTracked', 'closeForce'] as const;
const EXCEL_REOPEN_MODE_VALUES = ['lastState', 'readOnly', 'readWrite'] as const;
export type ExcelCoordinationMode = (typeof EXCEL_COORDINATION_MODE_VALUES)[number];
export type ExcelReopenMode = (typeof EXCEL_REOPEN_MODE_VALUES)[number];

function normalizeExcelCoordinationMode(value: unknown): ExcelCoordinationMode {
    return (EXCEL_COORDINATION_MODE_VALUES as readonly string[]).includes(value as string)
        ? (value as ExcelCoordinationMode)
        : 'block';
}

function normalizeExcelReopenMode(value: unknown): ExcelReopenMode {
    return (EXCEL_REOPEN_MODE_VALUES as readonly string[]).includes(value as string)
        ? (value as ExcelReopenMode)
        : 'lastState';
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
        webviewCard: {
            section: 'runtime',
            label: 'Python Path',
            description: 'Full path to the Python interpreter XLIDE runs its backend with. Leave blank to auto-detect Python on your PATH.',
            control: { kind: 'text' },
        },
    },
    'excelIntegration.coordinationMode': {
        defaultValue: (): ExcelCoordinationMode => 'block',
        normalize: normalizeExcelCoordinationMode,
        validate: (values, problems, key) => expectEnum(values, problems, key, EXCEL_COORDINATION_MODE_VALUES),
        manifest: { type: 'string', enum: EXCEL_COORDINATION_MODE_VALUES },
        webviewCard: {
            section: 'excel',
            label: 'When a Module is Blocked From Saving by Excel',
            description: 'What XLIDE does when Excel holds the workbook open for editing, which locks the file so a save, add, rename, delete, or F5 cannot write it. Block (default, safest): refuse and ask you to close it in Excel. Close Tracked: gracefully close a workbook XLIDE opened, then proceed. Close Force: close it in any Excel, force-quitting Excel if needed (unsafe; can lose unsaved work in other workbooks).',
            control: { kind: 'enum', values: EXCEL_COORDINATION_MODE_VALUES },
        },
    },
    'excelIntegration.trackOpenedWorkbooks': {
        defaultValue: () => true,
        normalize: normalizeBoolean(true),
        validate: expectBoolean,
        manifest: { type: 'boolean' },
        webviewCard: {
            section: 'excel',
            label: 'Close Only Workbooks XLIDE Opened',
            description: 'When the mode is "Close Tracked", only close workbooks XLIDE itself opened in Excel. Turn off to close a matching workbook in any running Excel, including ones you opened by hand. Ignored for Block and Close Force.',
            control: { kind: 'boolean' },
        },
    },
    'excelIntegration.reopenAfterClose': {
        defaultValue: () => true,
        normalize: normalizeBoolean(true),
        validate: expectBoolean,
        manifest: { type: 'boolean' },
        webviewCard: {
            section: 'excel',
            label: 'Reopen After Close',
            description: 'After XLIDE closes a workbook in Excel to write to it (a save, add, rename, or delete under a close mode), reopen it afterward so your Excel view is restored. Turn off to leave it closed until you reopen it yourself.',
            control: { kind: 'boolean' },
        },
    },
    'excelIntegration.reopenMode': {
        defaultValue: (): ExcelReopenMode => 'lastState',
        normalize: normalizeExcelReopenMode,
        validate: (values, problems, key) => expectEnum(values, problems, key, EXCEL_REOPEN_MODE_VALUES),
        manifest: { type: 'string', enum: EXCEL_REOPEN_MODE_VALUES },
        webviewCard: {
            section: 'excel',
            label: 'Reopen As',
            description: 'How XLIDE reopens a workbook it closed (when "Reopen After Close" is on). Last State (default): put it back the way it was, so read-only stays read-only and editable stays editable. Read-Only: always reopen read-only (keeps the file unlocked for your next save). Read-Write: reopen for editing in Excel (re-locks the file, so the next save closes it again).',
            control: { kind: 'enum', values: EXCEL_REOPEN_MODE_VALUES },
        },
    },
    'excelIntegration.reopenReadOnlyAfterSave': {
        defaultValue: () => false,
        normalize: normalizeBoolean(false),
        validate: expectBoolean,
        manifest: { type: 'boolean' },
        webviewCard: {
            section: 'excel',
            label: 'Reopen Read-Only Workbook After Module Save',
            description: 'A workbook open read-only in Excel does not lock the file, so XLIDE\'s save succeeds, but Excel keeps showing its older copy. Turn this on to silently close and reopen the read-only workbook after each save so Excel matches the saved file. Only acts when the workbook is actually open read-only; never reopens one you closed or one open for editing.',
            control: { kind: 'boolean' },
        },
    },
    'attachToRunningExcel': {
        defaultValue: () => true,
        normalize: normalizeBoolean(true),
        validate: expectBoolean,
        manifest: { type: 'boolean' },
        webviewCard: {
            section: 'excel',
            label: 'Attach To Running Excel',
            description: 'When opening a workbook or running a macro, reuse a running Excel instance and an already-open copy of the workbook before launching a fresh one. Most users keep this on.',
            control: { kind: 'boolean' },
        },
    },
    'diagnostics.enabled': {
        defaultValue: () => true,
        normalize: normalizeBoolean(true),
        validate: expectBoolean,
        manifest: { type: 'boolean' },
        webviewCard: {
            section: 'editor',
            label: 'Diagnostics Enabled',
            description: 'Show XLIDE\'s VBA diagnostics (red/yellow squiggles) in the editor. Turn off to silence all analysis warnings and errors.',
            control: { kind: 'boolean' },
        },
    },
    'editor.blockLayout': {
        defaultValue: () => DEFAULT_VBA_SMART_BLOCK_LAYOUT,
        normalize: normalizeSmartBlockLayout,
        validate: (values, problems, key) => expectEnum(values, problems, key, BLOCK_LAYOUT_VALUES),
        manifest: { type: 'string', enum: BLOCK_LAYOUT_VALUES },
        webviewCard: {
            section: 'editor',
            label: 'Editor Block Layout',
            description: 'Spacing of smart code blocks in the editor. Comfy adds breathing room; Compact is denser.',
            control: { kind: 'enum', values: BLOCK_LAYOUT_VALUES },
        },
    },
    'editor.continueCommentOnNewline': {
        defaultValue: () => true,
        normalize: normalizeBoolean(true),
        validate: expectBoolean,
        manifest: { type: 'boolean' },
        webviewCard: {
            section: 'editor',
            label: 'Continue Comment On New Line',
            description: "When the line above is a VBA comment (starts with an apostrophe), pressing Enter begins the new line with an apostrophe to continue the comment. Turn off for a plain new line with no apostrophe.",
            control: { kind: 'boolean' },
        },
    },
    'editor.mirrorCommentSpacing': {
        defaultValue: () => true,
        normalize: normalizeBoolean(true),
        validate: expectBoolean,
        manifest: { type: 'boolean' },
        webviewCard: {
            section: 'editor',
            label: 'Mirror Comment Spacing',
            description: "When continuing a comment, copy the run of spaces that follows the apostrophe on the line above so the text lines up. Turn off to insert just the apostrophe with no space. Only applies when 'Continue Comment On New Line' is on.",
            control: { kind: 'boolean' },
        },
    },
    'docs.enabled': {
        defaultValue: () => true,
        normalize: normalizeBoolean(true),
        validate: expectBoolean,
        manifest: { type: 'boolean' },
        webviewCard: {
            section: 'docs',
            label: 'Docs Enabled',
            description: 'Show hover documentation pulled from your .vbref.xml reference files alongside VBA symbols.',
            control: { kind: 'boolean' },
        },
    },
    'docs.metadataGlob': {
        defaultValue: () => DEFAULT_DOC_METADATA_GLOB,
        normalize: (value) => normalizeNonEmptyString(value, DEFAULT_DOC_METADATA_GLOB),
        validate: expectString,
        manifest: { type: 'string' },
        webviewCard: {
            section: 'docs',
            label: 'Docs Metadata Glob',
            description: 'Glob that locates your VBA documentation files (default **/*.vbref.xml). Used to attach hover docs to symbols.',
            control: { kind: 'text' },
        },
    },
    'analysis.visibleSeverities': {
        defaultValue: () => [...ANALYSIS_SEVERITIES],
        normalize: normalizeAnalysisVisibleSeverities,
        validate: (values, problems, key) => expectStringArrayEnum(values, problems, key, ANALYSIS_SEVERITIES),
        manifest: { type: 'array', items: { type: 'string', enum: ANALYSIS_SEVERITIES } },
        webviewCard: {
            section: 'analysis',
            label: 'Visible Severities',
            description: 'Which diagnostic severities XLIDE shows. Unchecked severities are hidden from the editor and Problems panel.',
            control: { kind: 'severityFilter' },
        },
    },
    'analysis.untrackedRules': {
        defaultValue: () => [],
        normalize: normalizeKnownAnalysisRuleCodes,
        validate: expectKnownAnalysisRuleCodeArray,
        manifest: { type: 'array', items: { type: 'string' } },
        webviewCard: {
            section: 'analysis',
            label: 'Globally Untracked Rules',
            description: 'Analysis rules to disable everywhere, across all workbooks. Use this to permanently silence a rule you never want.',
            control: { kind: 'rulePicker' },
        },
    },
    'analysis.ruleSeverityOverrides': {
        defaultValue: () => ({}),
        normalize: normalizeAnalysisRuleSeverityOverrides,
        validate: expectRuleSeverityOverrides,
        manifest: { type: 'object', additionalProperties: { type: 'string', enum: RULE_SEVERITY_OVERRIDE_VALUES } },
        webviewCard: {
            section: 'analysis',
            label: 'Rule Severity Overrides',
            description: 'Force specific rules to a chosen severity (e.g. downgrade a rule to a warning) globally, regardless of their defaults.',
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

function xlideExcelCoordinationModeFromConfig(config: vscode.WorkspaceConfiguration) {
    return xlideGlobalSettingFromConfig(config, 'excelIntegration.coordinationMode');
}

function xlideExcelTrackOpenedWorkbooksFromConfig(config: vscode.WorkspaceConfiguration) {
    return xlideGlobalSettingFromConfig(config, 'excelIntegration.trackOpenedWorkbooks');
}

function xlideExcelReopenAfterCloseFromConfig(config: vscode.WorkspaceConfiguration) {
    return xlideGlobalSettingFromConfig(config, 'excelIntegration.reopenAfterClose');
}

function xlideExcelReopenModeFromConfig(config: vscode.WorkspaceConfiguration) {
    return xlideGlobalSettingFromConfig(config, 'excelIntegration.reopenMode');
}

function xlideExcelReopenReadOnlyAfterSaveFromConfig(config: vscode.WorkspaceConfiguration) {
    return xlideGlobalSettingFromConfig(config, 'excelIntegration.reopenReadOnlyAfterSave');
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

function xlideEditorContinueCommentOnNewlineFromConfig(config: vscode.WorkspaceConfiguration) {
    return xlideGlobalSettingFromConfig(config, 'editor.continueCommentOnNewline');
}

function xlideEditorMirrorCommentSpacingFromConfig(config: vscode.WorkspaceConfiguration) {
    return xlideGlobalSettingFromConfig(config, 'editor.mirrorCommentSpacing');
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
    xlideExcelCoordinationModeFromConfig,
    xlideExcelTrackOpenedWorkbooksFromConfig,
    xlideExcelReopenAfterCloseFromConfig,
    xlideExcelReopenModeFromConfig,
    xlideExcelReopenReadOnlyAfterSaveFromConfig,
    xlideDiagnosticsEnabledFromConfig,
    xlideDocsEnabledFromConfig,
    xlideDocsMetadataGlobFromConfig,
    xlideEditorBlockLayoutFromConfig,
    xlideEditorContinueCommentOnNewlineFromConfig,
    xlideEditorMirrorCommentSpacingFromConfig,
    xlideGlobalSettingCards,
    xlideGlobalSettingManifest,
    xlidePerformanceTraceFromConfig,
    xlidePythonPathFromConfig,
};
