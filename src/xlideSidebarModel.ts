import type { ResolvedXlideGlobalSetting } from './globalSettings';
import type { EffectiveWorkbookAnalysisSettings } from './workbookAnalysisSettings';
import type { EffectiveWorkbookModuleSyncSettings } from './workbookModuleSyncSettings';

type XlideSidebarNodeKind = 'section' | 'status' | 'setting' | 'action';
type XlideSidebarStatus = 'pass' | 'warn' | 'fail' | 'unknown';
type XlideSidebarSettingSource = ResolvedXlideGlobalSetting<unknown>['source'] | 'workbook' | 'session' | 'missing';

interface XlideSidebarActiveWorkbook {
    label: string;
    filePath: string;
    settingsPath: string;
    selectionSource: 'activeEditor' | 'singleWorkbook';
    settingsState: 'missing' | 'valid' | 'invalid';
    settingsMessage?: string;
    analysisSettings?: EffectiveWorkbookAnalysisSettings;
    moduleSyncSettings?: EffectiveWorkbookModuleSyncSettings;
}

interface XlideSidebarCommand {
    command: string;
    title: string;
    arguments?: unknown[];
}

interface XlideSidebarNode {
    id: string;
    kind: XlideSidebarNodeKind;
    label: string;
    description?: string;
    tooltip?: string;
    icon?: string;
    status?: XlideSidebarStatus;
    command?: XlideSidebarCommand;
    children?: XlideSidebarNode[];
}

interface XlideSidebarModelInput {
    globalSettings: readonly ResolvedXlideGlobalSetting<unknown>[];
    hasWorkspace: boolean;
    workbookCount: number;
    activeWorkbook?: XlideSidebarActiveWorkbook;
}

const GLOBAL_SETTING_LABELS: Record<string, string> = {
    'xlide.analysis.ruleSeverityOverrides': 'Rule severities',
    'xlide.analysis.untrackedRules': 'Untracked rules',
    'xlide.analysis.visibleSeverities': 'Visible severities',
    'xlide.attachToRunningExcel': 'Attach to running Excel',
    'xlide.diagnostics.enabled': 'Live diagnostics',
    'xlide.docs.enabled': 'Documentation hovers',
    'xlide.docs.metadataGlob': 'Doc metadata glob',
    'xlide.editor.blockLayout': 'Block layout',
    'xlide.pythonPath': 'Python executable',
};

function buildXlideSidebarModel(input: XlideSidebarModelInput): XlideSidebarNode[] {
    const workbookArg = input.activeWorkbook
        ? {
            kind: 'xlsm',
            label: input.activeWorkbook.label,
            filePath: input.activeWorkbook.filePath,
        }
        : undefined;
    return [
        section('project', 'Project', [
            statusNode(
                'project.workspace',
                'Workspace',
                input.hasWorkspace ? 'Open' : 'No folder',
                input.hasWorkspace ? 'pass' : 'warn',
                input.hasWorkspace
                    ? 'VS Code has a workspace folder open.'
                    : 'Open a workspace folder so XLIDE can discover workbooks.',
            ),
            statusNode(
                'project.workbooks',
                'Workbook discovery',
                workbookDiscoveryDescription(input.workbookCount),
                input.workbookCount > 0 ? 'pass' : 'unknown',
                input.workbookCount > 0
                    ? `${input.workbookCount} Excel VBA workbook(s) were found in the workspace.`
                    : 'No .xlsm, .xlsb, or .xlam workbooks were found in the workspace yet.',
            ),
            activeWorkbookNode(input.activeWorkbook),
            workbookSettingsStatusNode(input.activeWorkbook),
            actionNode(
                'project.revealExplorer',
                'Reveal workbook tree',
                'Explorer',
                'xlide.explorer.focus',
                'Focus the XLIDE workbook tree in the VS Code Explorer.',
                'list-tree',
            ),
            actionNode(
                'project.refreshExplorer',
                'Refresh workbooks',
                undefined,
                'xlide.refreshExplorer',
                'Refresh the XLIDE workbook tree.',
                'refresh',
            ),
        ]),
        section('actions', 'Actions', [
            actionNode(
                'actions.analyzeCurrentModule',
                'Analyze current module',
                undefined,
                'xlide.analyzeCurrentModule',
                'Analyze the active VBA module using the same engine as workbook analysis.',
                'check',
            ),
            actionNode(
                'actions.analyzeWorkbook',
                'Analyze active workbook',
                undefined,
                'xlide.analyzeWorkbook',
                'Analyze the active workbook. Open a workbook module first if no workbook node is selected.',
                'checklist',
                workbookArg ? [workbookArg] : [],
            ),
            actionNode(
                'actions.importModules',
                'Import modules',
                undefined,
                'xlide.importModulesFromFolder',
                'Open the module import diff GUI for the active workbook.',
                'import',
                workbookArg ? [workbookArg] : [],
            ),
            actionNode(
                'actions.exportModules',
                'Export modules',
                undefined,
                'xlide.exportModulesToFolder',
                'Open the module export diff GUI for the active workbook.',
                'export',
                workbookArg ? [workbookArg] : [],
            ),
            actionNode(
                'actions.validateWorkbook',
                'Validate VBA project',
                undefined,
                'xlide.validateWorkbook',
                'Validate the active workbook VBA project structure.',
                'verified',
                workbookArg ? [workbookArg] : [],
            ),
            actionNode(
                'actions.openWorkbook',
                'Open workbook in Excel',
                undefined,
                'xlide.openWorkbook',
                'Open the active workbook in Excel.',
                'file-excel',
                workbookArg ? [workbookArg] : [],
            ),
        ]),
        workbookConfigurationSection(input.activeWorkbook),
        section('configuration', 'Configuration', [
            ...input.globalSettings.map((setting) => settingNode(setting)),
            actionNode(
                'configuration.openSettings',
                'Open global settings',
                'VS Code',
                'workbench.action.openSettings',
                'Open XLIDE global/editor settings in VS Code Settings.',
                'settings-gear',
                ['@ext:WilliamSmithE.xlide'],
            ),
        ]),
        section('support', 'Support', [
            actionNode(
                'support.setup',
                'Install Python dependencies',
                undefined,
                'xlide.setup',
                'Install or repair the Python dependencies used by the XLIDE backend.',
                'cloud-download',
            ),
            actionNode(
                'support.copyDiagnostics',
                'Copy diagnostics',
                undefined,
                'xlide.copyDiagnostics',
                'Copy a redacted diagnostic snapshot to the clipboard.',
                'clippy',
            ),
            actionNode(
                'support.exportBundle',
                'Export support bundle',
                undefined,
                'xlide.exportSupportBundle',
                'Export a redacted support bundle for troubleshooting.',
                'archive',
            ),
        ]),
    ];
}

function section(id: string, label: string, children: XlideSidebarNode[]): XlideSidebarNode {
    return { id, kind: 'section', label, icon: 'folder', children };
}

function statusNode(
    id: string,
    label: string,
    description: string,
    status: XlideSidebarStatus,
    tooltip: string,
    icon = statusIcon(status),
): XlideSidebarNode {
    return { id, kind: 'status', label, description, status, tooltip, icon };
}

function activeWorkbookNode(workbook: XlideSidebarActiveWorkbook | undefined): XlideSidebarNode {
    if (!workbook) {
        return statusNode(
            'project.activeWorkbook',
            'Active workbook',
            'None selected',
            'unknown',
            'Open an XLIDE VBA module, or keep exactly one workbook in the workspace, to select a workbook context.',
        );
    }
    return statusNode(
        'project.activeWorkbook',
        'Active workbook',
        workbook.label,
        'pass',
        `${workbook.filePath}\nSelected from ${workbook.selectionSource === 'activeEditor' ? 'the active editor' : 'the only workbook in the workspace'}.`,
        'file-code',
    );
}

function workbookSettingsStatusNode(workbook: XlideSidebarActiveWorkbook | undefined): XlideSidebarNode {
    if (!workbook) {
        return statusNode(
            'project.workbookSettings',
            'Workbook settings',
            'No workbook',
            'unknown',
            'Workbook-scoped settings are shown after XLIDE has an active workbook context.',
        );
    }
    if (workbook.settingsState === 'invalid') {
        return statusNode(
            'project.workbookSettings',
            'Workbook settings',
            'Invalid',
            'fail',
            workbook.settingsMessage ?? workbook.settingsPath,
        );
    }
    if (workbook.settingsState === 'missing') {
        return statusNode(
            'project.workbookSettings',
            'Workbook settings',
            'Using defaults',
            'unknown',
            `${workbook.settingsPath}\nNo workbook sidecar exists yet, so workbook-facing GUIs use global/editor defaults.`,
        );
    }
    return statusNode(
        'project.workbookSettings',
        'Workbook settings',
        'Sidecar loaded',
        'pass',
        workbook.settingsPath,
        'json',
    );
}

function workbookConfigurationSection(workbook: XlideSidebarActiveWorkbook | undefined): XlideSidebarNode {
    if (!workbook) {
        return section('workbookConfiguration', 'Workbook Configuration', [
            statusNode(
                'workbookConfiguration.none',
                'Workbook context',
                'None selected',
                'unknown',
                'Open an XLIDE VBA module or select a workspace with one workbook to see workbook-scoped settings.',
            ),
        ]);
    }
    if (workbook.settingsState === 'invalid') {
        return section('workbookConfiguration', 'Workbook Configuration', [
            workbookSettingsStatusNode(workbook),
        ]);
    }

    const analysis = workbook.analysisSettings;
    const sync = workbook.moduleSyncSettings;
    const children: XlideSidebarNode[] = [
        workbookSettingsStatusNode(workbook),
    ];
    if (workbook.settingsState === 'valid') {
        children.push(actionNode(
            'workbookConfiguration.openSidecar',
            'Open workbook settings',
            'Sidecar',
            'xlide.openWorkbookSettings',
            'Open the workbook-scoped XLIDE settings sidecar.',
            'json',
            [workbook.settingsPath],
        ));
    }
    if (sync) {
        children.push(
            settingLikeNode(
                'workbookConfiguration.exportFolder',
                'Export folder',
                sync.folderPath ?? 'Not configured',
                sync.folderPathSource,
                sync.folderPath ? 'folder' : 'warning',
            ),
            settingLikeNode(
                'workbookConfiguration.exportMode',
                'Export mode',
                exportModeLabel(sync.exportMode),
                sync.exportModeSource,
            ),
            settingLikeNode(
                'workbookConfiguration.importMode',
                'Import mode',
                importModeLabel(sync.importMode),
                sync.importModeSource,
            ),
        );
    }
    if (analysis) {
        children.push(
            settingLikeNode(
                'workbookConfiguration.analysisVisibleSeverities',
                'Analysis severities',
                settingDescription(analysis.visibleSeverities),
                analysis.visibleSeveritiesSource,
                'list-selection',
            ),
            settingLikeNode(
                'workbookConfiguration.analysisUntrackedRules',
                'Untracked rules',
                settingDescription(analysis.untrackedRules),
                analysis.untrackedRulesSource,
                'list-unordered',
            ),
            settingLikeNode(
                'workbookConfiguration.analysisRuleSeverities',
                'Rule severities',
                settingDescription(analysis.ruleSeverityOverrides),
                analysis.ruleSeverityOverridesSource,
                'json',
            ),
        );
    }
    return section('workbookConfiguration', 'Workbook Configuration', children);
}

function settingNode(setting: ResolvedXlideGlobalSetting<unknown>): XlideSidebarNode {
    const key = setting.key;
    const label = GLOBAL_SETTING_LABELS[key] ?? key.replace(/^xlide\./, '');
    const description = `${settingDescription(setting.value)} (${sourceLabel(setting.source)})`;
    return {
        id: `configuration.${key}`,
        kind: 'setting',
        label,
        description,
        tooltip: `${key}\n${description}`,
        icon: settingIcon(setting.value),
        command: {
            command: 'workbench.action.openSettings',
            title: 'Open Setting',
            arguments: [key],
        },
    };
}

function settingLikeNode(
    id: string,
    label: string,
    value: string,
    source: XlideSidebarSettingSource,
    icon = 'settings',
): XlideSidebarNode {
    const description = `${value} (${sourceLabel(source)})`;
    return {
        id,
        kind: 'setting',
        label,
        description,
        tooltip: description,
        icon,
    };
}

function actionNode(
    id: string,
    label: string,
    description: string | undefined,
    command: string,
    tooltip: string,
    icon: string,
    args: unknown[] = [],
): XlideSidebarNode {
    return {
        id,
        kind: 'action',
        label,
        description,
        tooltip,
        icon,
        command: { command, title: label, arguments: args },
    };
}

function workbookDiscoveryDescription(workbookCount: number): string {
    if (workbookCount === 0) {
        return 'None found';
    }
    if (workbookCount === 1) {
        return '1 workbook';
    }
    return `${workbookCount} workbooks`;
}

function settingDescription(value: unknown): string {
    if (typeof value === 'boolean') {
        return value ? 'Enabled' : 'Disabled';
    }
    if (typeof value === 'string') {
        return value.length > 0 ? value : 'From PATH';
    }
    if (Array.isArray(value)) {
        return value.length > 0 ? value.join(', ') : 'None';
    }
    if (value && typeof value === 'object') {
        const count = Object.keys(value).length;
        return count === 1 ? '1 override' : `${count} overrides`;
    }
    if (value === undefined || value === null) {
        return 'Not set';
    }
    return String(value);
}

function sourceLabel(source: XlideSidebarSettingSource): string {
    switch (source) {
        case 'workbook':
            return 'Workbook';
        case 'session':
            return 'Session';
        case 'missing':
            return 'Not set';
        case 'machine':
            return 'VS Code';
        case 'default':
            return 'Default';
        default:
            return 'Unknown';
    }
}

function statusIcon(status: XlideSidebarStatus): string {
    switch (status) {
        case 'pass':
            return 'pass';
        case 'warn':
            return 'warning';
        case 'fail':
            return 'error';
        default:
            return 'question';
    }
}

function settingIcon(value: unknown): string {
    if (typeof value === 'boolean') {
        return value ? 'check' : 'circle-slash';
    }
    if (Array.isArray(value)) {
        return 'list-selection';
    }
    if (value && typeof value === 'object') {
        return 'json';
    }
    return 'settings';
}

function exportModeLabel(mode: string): string {
    return mode === 'trueUp' ? 'Export All + Delete Missing' : 'Export All';
}

function importModeLabel(mode: string): string {
    return mode === 'trueUpStandardClass' ? 'Import/Update + Delete Missing' : 'Import/Update';
}

export {
    buildXlideSidebarModel,
    settingDescription,
    sourceLabel,
    type XlideSidebarActiveWorkbook,
    type XlideSidebarCommand,
    type XlideSidebarModelInput,
    type XlideSidebarNode,
    type XlideSidebarNodeKind,
    type XlideSidebarStatus,
};
