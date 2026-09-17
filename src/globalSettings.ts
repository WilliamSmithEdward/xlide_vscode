import type * as vscode from 'vscode';
import {
    ANALYSIS_SEVERITIES,
    normalizeKnownAnalysisRuleCodes,
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
import { normalizeDiagnosticRuleCode } from './analyzer/diagnostics/ruleMetadata';

type XlideGlobalSettingSeverity = 'warning';

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
    'officeIntegration.attachToRunning': boolean;
    'officeIntegration.coordinationMode': OfficeCoordinationMode;
    'officeIntegration.trackOpenedFiles': boolean;
    'officeIntegration.reopenAfterClose': boolean;
    'officeIntegration.reopenMode': OfficeReopenMode;
    'officeIntegration.reopenReadOnlyAfterSave': boolean;
    'formRun.injectShowMacro': FormRunInjectShowMacro;
    'agent.showWriteDiffs': boolean;
    'diagnostics.enabled': boolean;
    'editor.blockLayout': VbaSmartBlockLayout;
    'editor.continueCommentOnNewline': boolean;
    'editor.mirrorCommentSpacing': boolean;
    'explorer.autoExpandCollapse': boolean;
    'explorer.view': XlideExplorerView;
    'docs.enabled': boolean;
    'docs.metadataGlob': string;
    'analysis.visibleSeverities': AnalysisSeverityFilter[];
    'analysis.ignoreFilesOutsideTree': boolean;
    'analysis.untrackedRules': string[];
    'analysis.ruleSeverityOverrides': AnalysisRuleSeverityOverrides;
    'performance.trace': boolean;
}

type XlideGlobalSettingKey = keyof XlideGlobalSettingValues;
type XlideGlobalSettingSection = 'office' | 'editor' | 'docs' | 'analysis';

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
    /**
     * The name this setting had before it was renamed. A value still stored
     * under it is honored until the setting is given one under its new name,
     * and the old name stays contributed, marked deprecated, so VS Code does
     * not flag a user's existing entry as unknown.
     */
    legacyKey?: string;
}

const OFFICE_COORDINATION_MODE_VALUES = ['block', 'closeTracked', 'closeForce'] as const;
const OFFICE_REOPEN_MODE_VALUES = ['lastState', 'readOnly', 'readWrite'] as const;
export type OfficeCoordinationMode = (typeof OFFICE_COORDINATION_MODE_VALUES)[number];
export type OfficeReopenMode = (typeof OFFICE_REOPEN_MODE_VALUES)[number];

function normalizeOfficeCoordinationMode(value: unknown): OfficeCoordinationMode {
    return (OFFICE_COORDINATION_MODE_VALUES as readonly string[]).includes(value as string)
        ? (value as OfficeCoordinationMode)
        : 'block';
}

function normalizeOfficeReopenMode(value: unknown): OfficeReopenMode {
    return (OFFICE_REOPEN_MODE_VALUES as readonly string[]).includes(value as string)
        ? (value as OfficeReopenMode)
        : 'lastState';
}

const FORM_RUN_INJECT_SHOW_MACRO_VALUES = ['ask', 'always', 'never'] as const;

export type FormRunInjectShowMacro = (typeof FORM_RUN_INJECT_SHOW_MACRO_VALUES)[number];

function normalizeFormRunInjectShowMacro(value: unknown): FormRunInjectShowMacro {
    return (FORM_RUN_INJECT_SHOW_MACRO_VALUES as readonly string[]).includes(value as string)
        ? (value as FormRunInjectShowMacro)
        : 'ask';
}

const BLOCK_LAYOUT_VALUES = ['comfy', 'compact'] as const;
const EXPLORER_VIEW_VALUES = ['tree', 'folders'] as const;
/** Which layout the XLIDE explorer draws a project's modules in. */
export type XlideExplorerView = (typeof EXPLORER_VIEW_VALUES)[number];
const RULE_SEVERITY_OVERRIDE_VALUES = ['off', 'warning'] as const;
const DEFAULT_DOC_METADATA_GLOB = '**/*.vbref.xml';

// Single source of truth for every xlide.* global setting, declared in
// settings-page display order. The package.json contributes.configuration
// block is asserted against xlideGlobalSettingManifest() by
// tests/globalSettingsManifest.test.ts.
const XLIDE_GLOBAL_SETTINGS: {
    [K in XlideGlobalSettingKey]: XlideGlobalSettingSchema<XlideGlobalSettingValues[K]>;
} = {
    'officeIntegration.coordinationMode': {
        defaultValue: (): OfficeCoordinationMode => 'block',
        normalize: normalizeOfficeCoordinationMode,
        validate: (values, problems, key) => expectEnum(values, problems, key, OFFICE_COORDINATION_MODE_VALUES),
        manifest: { type: 'string', enum: OFFICE_COORDINATION_MODE_VALUES },
        legacyKey: 'excelIntegration.coordinationMode',
        webviewCard: {
            section: 'office',
            label: 'When a File is Open in Its Application',
            description: 'What XLIDE does when Excel, Word, PowerPoint or Access holds the file open for editing, which locks it so a save, add, rename, delete, or F5 cannot write it. Block (default, safest): refuse and ask you to close it in its application. Close Tracked: gracefully close a file XLIDE opened, then proceed. Close Force: close it in any running instance, force-quitting the application if needed (unsafe; can lose unsaved work in its other open files). Under every mode XLIDE closes and reopens a read-only copy it opened itself, which holds nothing to lose.',
            control: { kind: 'enum', values: OFFICE_COORDINATION_MODE_VALUES },
        },
    },
    'officeIntegration.trackOpenedFiles': {
        defaultValue: () => true,
        normalize: normalizeBoolean(true),
        validate: expectBoolean,
        manifest: { type: 'boolean' },
        legacyKey: 'excelIntegration.trackOpenedWorkbooks',
        webviewCard: {
            section: 'office',
            label: 'Close Only Files XLIDE Opened',
            description: 'When the mode is "Close Tracked", only close files XLIDE itself opened in their application. Turn off to close a matching file in any running instance, including ones you opened by hand. Ignored for Block and Close Force.',
            control: { kind: 'boolean' },
        },
    },
    'officeIntegration.reopenAfterClose': {
        defaultValue: () => true,
        normalize: normalizeBoolean(true),
        validate: expectBoolean,
        manifest: { type: 'boolean' },
        legacyKey: 'excelIntegration.reopenAfterClose',
        webviewCard: {
            section: 'office',
            label: 'Reopen After Close',
            description: 'After XLIDE closes a file in its application to write to it (a save, add, rename, or delete), reopen it afterward so your view is restored. Turn off to leave it closed until you reopen it yourself.',
            control: { kind: 'boolean' },
        },
    },
    'officeIntegration.reopenMode': {
        defaultValue: (): OfficeReopenMode => 'lastState',
        normalize: normalizeOfficeReopenMode,
        validate: (values, problems, key) => expectEnum(values, problems, key, OFFICE_REOPEN_MODE_VALUES),
        manifest: { type: 'string', enum: OFFICE_REOPEN_MODE_VALUES },
        legacyKey: 'excelIntegration.reopenMode',
        webviewCard: {
            section: 'office',
            label: 'Reopen As',
            description: 'How XLIDE reopens a file it closed (when "Reopen After Close" is on). Last State (default): put it back the way it was, so read-only stays read-only and editable stays editable. Read-Only: always reopen read-only. Read-Write: reopen for editing, which locks the file again, so the next save closes it again. Access has no read-only open, so a database always reopens for editing.',
            control: { kind: 'enum', values: OFFICE_REOPEN_MODE_VALUES },
        },
    },
    'officeIntegration.reopenReadOnlyAfterSave': {
        defaultValue: () => false,
        normalize: normalizeBoolean(false),
        validate: expectBoolean,
        manifest: { type: 'boolean' },
        legacyKey: 'excelIntegration.reopenReadOnlyAfterSave',
        webviewCard: {
            section: 'office',
            label: 'Refresh a Read-Only Copy After Module Save',
            description: 'A workbook open read-only in Excel does not lock the file, so XLIDE\'s save succeeds, but Excel keeps showing its older copy. Turn this on to silently close and reopen the read-only copy after each save so the application matches the saved file. Only acts when the file is actually open read-only; never reopens one you closed or one open for editing. Word and PowerPoint lock the file even when it is read-only, so there XLIDE closes and reopens its own read-only copy around every save regardless of this setting.',
            control: { kind: 'boolean' },
        },
    },
    'formRun.injectShowMacro': {
        defaultValue: (): FormRunInjectShowMacro => 'ask',
        normalize: normalizeFormRunInjectShowMacro,
        validate: (values, problems, key) => expectEnum(values, problems, key, FORM_RUN_INJECT_SHOW_MACRO_VALUES),
        manifest: { type: 'string', enum: FORM_RUN_INJECT_SHOW_MACRO_VALUES },
        webviewCard: {
            section: 'office',
            label: 'Run Form (F5) Show Macro',
            description: 'When F5 launches a UserForm in an Excel, Word or PowerPoint file, XLIDE can add a small launcher macro to the file - one sub per form, all in module XlideRun - and run it so the form opens immediately, the way F5 in the VBE does. Ask (default): confirm the first time a form needs its launcher; once that sub exists F5 just runs it, with no prompt. Always: add and show without asking. Never: F5 just opens the file. The subs stay in the file and are safe to delete. An Access form or report needs no launcher: Access opens it by name, so nothing is added to the database.',
            control: { kind: 'enum', values: FORM_RUN_INJECT_SHOW_MACRO_VALUES },
        },
    },
    'officeIntegration.attachToRunning': {
        defaultValue: () => true,
        normalize: normalizeBoolean(true),
        validate: expectBoolean,
        manifest: { type: 'boolean' },
        legacyKey: 'attachToRunningExcel',
        webviewCard: {
            section: 'office',
            label: 'Attach To Running Application',
            description: 'When opening a file or running a macro, reuse a running instance of the file\'s application (Excel, Word or Access) and an already-open copy of the file before launching a fresh one. PowerPoint only ever runs one instance. Most users keep this on: turned off, every open and F5 starts another instance, and in Word and Access each one keeps the file locked until you close it.',
            control: { kind: 'boolean' },
        },
    },
    'agent.showWriteDiffs': {
        defaultValue: () => true,
        normalize: normalizeBoolean(true),
        validate: expectBoolean,
        manifest: { type: 'boolean' },
        webviewCard: {
            section: 'editor',
            label: 'Review Agent Writes',
            description: 'When an AI agent writes a VBA module through the XLIDE tools, open a before/after diff and badge the module in the XLIDE tree with Keep / Revert actions. Agent tool writes never pass through the editor, so without this no diff appears anywhere.',
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
    'explorer.autoExpandCollapse': {
        defaultValue: () => true,
        normalize: normalizeBoolean(true),
        validate: expectBoolean,
        manifest: { type: 'boolean' },
        webviewCard: {
            section: 'editor',
            label: 'Auto Expand And Collapse Explorer Tree',
            description: "Automatically reveal the active module in the XLIDE explorer and collapse the others as you switch editor tabs (a one-file, one-module accordion). Turn off to leave the tree as you arrange it - switching tabs and expanding nodes will not auto-collapse anything.",
            control: { kind: 'boolean' },
        },
    },
    'explorer.view': {
        defaultValue: () => 'tree' as XlideExplorerView,
        normalize: (value) => (value === 'folders' ? 'folders' : 'tree'),
        validate: (values, problems, key) => expectEnum(values, problems, key, EXPLORER_VIEW_VALUES),
        manifest: { type: 'string', enum: EXPLORER_VIEW_VALUES },
        webviewCard: {
            section: 'editor',
            label: 'Explorer Layout',
            description: "How the XLIDE explorer arranges a project's modules. Tree lists them flat. Folders groups them by the '@Folder(\"Name.Sub\")' comment in each module's declarations, the Rubberduck convention. The Tree and Folders buttons above the explorer switch it too.",
            control: { kind: 'enum', values: EXPLORER_VIEW_VALUES },
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
    'analysis.ignoreFilesOutsideTree': {
        defaultValue: () => true,
        normalize: normalizeBoolean(true),
        validate: expectBoolean,
        manifest: { type: 'boolean' },
        webviewCard: {
            section: 'analysis',
            label: 'Ignore Files Outside The XLIDE Tree',
            description: 'Analyze only the modules the XLIDE tree lists: the ones inside a workbook, document, presentation or database, and the files a VB6 project names. A .bas, .cls or .frm on disk that no project claims, such as an exported copy, is not analyzed, so its findings do not repeat the real module\'s in the Problems panel. Completion, hover and navigation still work in those files. Turn off to analyze them as standalone modules.',
            control: { kind: 'boolean' },
        },
    },
    'analysis.untrackedRules': {
        defaultValue: () => [],
        normalize: normalizeKnownAnalysisRuleCodes,
        validate: expectAnalysisRuleCodeArray,
        manifest: { type: 'array', items: { type: 'string' } },
        webviewCard: {
            section: 'analysis',
            label: 'Globally Untracked Rules',
            description: 'Analysis rules to disable everywhere, across all projects. Use this to permanently silence a rule you never want.',
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
        if (schema.legacyKey) {
            manifest[`xlide.${schema.legacyKey}`] = {
                ...schema.manifest,
                scope: 'machine',
                default: schema.defaultValue(),
                deprecationMessage: legacySettingDeprecationMessage(key),
            };
        }
    }
    return manifest;
}

/** What VS Code shows beside a setting still stored under its old name. */
function legacySettingDeprecationMessage(key: XlideGlobalSettingKey): string {
    return `Renamed to xlide.${key}. The value here still applies until that setting is given one.`;
}

/** The value a user explicitly stored for a key, at any scope. */
function explicitSettingValue(config: vscode.WorkspaceConfiguration, key: string): unknown {
    const inspected = typeof config.inspect === 'function' ? config.inspect<unknown>(key) : undefined;
    return inspected?.globalValue ?? inspected?.workspaceValue ?? inspected?.workspaceFolderValue;
}

function xlideGlobalSettingFromConfig<K extends XlideGlobalSettingKey>(
    config: vscode.WorkspaceConfiguration,
    key: K,
): ResolvedXlideGlobalSetting<XlideGlobalSettingValues[K]> {
    const schema = XLIDE_GLOBAL_SETTINGS[key];
    // A renamed setting keeps answering from its old name until the new one
    // is given a value, so an upgrade never silently drops a user's choice.
    const legacyValue = schema.legacyKey !== undefined && explicitSettingValue(config, key) === undefined
        ? explicitSettingValue(config, schema.legacyKey)
        : undefined;
    if (legacyValue !== undefined) {
        return { key: `xlide.${key}`, value: schema.normalize(legacyValue), source: 'machine' };
    }
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

function xlideOfficeAttachToRunningFromConfig(config: vscode.WorkspaceConfiguration) {
    return xlideGlobalSettingFromConfig(config, 'officeIntegration.attachToRunning');
}

function xlideOfficeCoordinationModeFromConfig(config: vscode.WorkspaceConfiguration) {
    return xlideGlobalSettingFromConfig(config, 'officeIntegration.coordinationMode');
}

function xlideOfficeTrackOpenedFilesFromConfig(config: vscode.WorkspaceConfiguration) {
    return xlideGlobalSettingFromConfig(config, 'officeIntegration.trackOpenedFiles');
}

function xlideOfficeReopenAfterCloseFromConfig(config: vscode.WorkspaceConfiguration) {
    return xlideGlobalSettingFromConfig(config, 'officeIntegration.reopenAfterClose');
}

function xlideOfficeReopenModeFromConfig(config: vscode.WorkspaceConfiguration) {
    return xlideGlobalSettingFromConfig(config, 'officeIntegration.reopenMode');
}

function xlideOfficeReopenReadOnlyAfterSaveFromConfig(config: vscode.WorkspaceConfiguration) {
    return xlideGlobalSettingFromConfig(config, 'officeIntegration.reopenReadOnlyAfterSave');
}

function xlideDiagnosticsEnabledFromConfig(config: vscode.WorkspaceConfiguration) {
    return xlideGlobalSettingFromConfig(config, 'diagnostics.enabled');
}

function xlideAnalysisIgnoreFilesOutsideTreeFromConfig(config: vscode.WorkspaceConfiguration) {
    return xlideGlobalSettingFromConfig(config, 'analysis.ignoreFilesOutsideTree');
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

function xlideExplorerAutoExpandCollapseFromConfig(config: vscode.WorkspaceConfiguration) {
    return xlideGlobalSettingFromConfig(config, 'explorer.autoExpandCollapse');
}

function xlideExplorerViewFromConfig(config: vscode.WorkspaceConfiguration) {
    return xlideGlobalSettingFromConfig(config, 'explorer.view');
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
        // An absent/unset value is always valid - the read path applies the
        // declared default. Only an explicitly-provided value is validated, so a
        // freshly-added setting never reports a problem before it is configured.
        if (values[key] === undefined) {
            continue;
        }
        XLIDE_GLOBAL_SETTINGS[key].validate(values, problems, key);
    }
    return problems;
}

function validateXlideGlobalSettingsFromConfig(
    config: vscode.WorkspaceConfiguration,
): XlideGlobalSettingsProblem[] {
    const snapshot: XlideGlobalSettingsSnapshot = {};
    for (const key of XLIDE_GLOBAL_SETTING_KEYS) {
        // Only validate a value the user explicitly set. A key that is unset (or
        // only carries its package.json default) must never be reported: on a
        // version upgrade, a newly-contributed key can briefly be absent from
        // the configuration registry, so config.get(key) returns undefined and
        // the read path applies the default silently. Validating those was the
        // upgrade "settings not set correctly" warning blast.
        const inspected = typeof config.inspect === 'function' ? config.inspect<unknown>(key) : undefined;
        const explicit = inspected?.globalValue ?? inspected?.workspaceValue ?? inspected?.workspaceFolderValue;
        if (explicit !== undefined) {
            snapshot[key] = explicit;
        }
    }
    return validateXlideGlobalSettingsValues(snapshot);
}

async function setXlideGlobalAnalysisRuleTracked(
    config: vscode.WorkspaceConfiguration,
    code: string | undefined,
    tracked: boolean,
): Promise<AnalysisRuleTrackingUpdate> {
    const normalized = normalizeDiagnosticRuleCode(code);
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
    const normalized = normalizeDiagnosticRuleCode(code);
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
    const normalized = normalizeDiagnosticRuleCode(code);
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
    let changed = inspect?.globalValue !== undefined;
    if (changed) {
        await config.update(key, undefined, true);
    }
    // A value left under the setting's old name would show through the reset.
    const legacyKey = XLIDE_GLOBAL_SETTINGS[key].legacyKey;
    if (legacyKey !== undefined && explicitSettingValue(config, legacyKey) !== undefined) {
        await config.update(legacyKey, undefined, true);
        changed = true;
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

function expectAnalysisRuleCodeArray(
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
    // A well-formed but unknown/renamed rule code is tolerated, not flagged: every
    // apply path drops unknown codes silently, so warning here would only nag the
    // user after a version renamed a code, with no way to tell which entry is stale.
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
        // A malformed value is non-fatal - XLIDE falls back to a safe default -
        // so surface it as a gentle warning, never a hard error.
        severity: 'warning',
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
    xlideOfficeAttachToRunningFromConfig,
    xlideOfficeCoordinationModeFromConfig,
    xlideOfficeTrackOpenedFilesFromConfig,
    xlideOfficeReopenAfterCloseFromConfig,
    xlideOfficeReopenModeFromConfig,
    xlideOfficeReopenReadOnlyAfterSaveFromConfig,
    xlideAnalysisIgnoreFilesOutsideTreeFromConfig,
    xlideDiagnosticsEnabledFromConfig,
    xlideDocsEnabledFromConfig,
    xlideDocsMetadataGlobFromConfig,
    xlideEditorBlockLayoutFromConfig,
    xlideEditorContinueCommentOnNewlineFromConfig,
    xlideEditorMirrorCommentSpacingFromConfig,
    xlideExplorerAutoExpandCollapseFromConfig,
    xlideExplorerViewFromConfig,
    xlideGlobalSettingCards,
    xlideGlobalSettingManifest,
    xlidePerformanceTraceFromConfig,
};
