type XlideSidebarNodeKind = 'section' | 'status' | 'action' | 'select';
type XlideSidebarStatus = 'pass' | 'warn' | 'fail' | 'unknown';
type XlideSidebarSetupAction = 'downloadPython' | 'setPythonPath';

interface XlideSidebarActiveWorkbook {
    label: string;
    filePath: string;
    settingsPath: string;
    selectionSource: 'activeEditor' | 'singleWorkbook' | 'sidebarSelection';
    settingsState: 'missing' | 'valid' | 'invalid';
    settingsMessage?: string;
}

interface XlideSidebarWorkbookChoice {
    label: string;
    filePath: string;
    description?: string;
}

interface XlideSidebarCommand {
    command: string;
    title: string;
    arguments?: unknown[];
    tooltip?: string;
    ctrlCommand?: string;
    ctrlTitle?: string;
    ctrlArguments?: unknown[];
    ctrlTooltip?: string;
}

interface XlideSidebarSelectOption {
    label: string;
    value: string;
    description?: string;
}

interface XlideSidebarDependencyStatus {
    status: XlideSidebarStatus;
    description: string;
    tooltip: string;
    action?: XlideSidebarSetupAction;
}

interface XlideSidebarSetupStatus {
    pythonExecutable: XlideSidebarDependencyStatus;
    pythonLibraries: XlideSidebarDependencyStatus;
}

interface XlideSidebarNode {
    id: string;
    kind: XlideSidebarNodeKind;
    label: string;
    description?: string;
    tooltip?: string;
    icon?: string;
    status?: XlideSidebarStatus;
    disabled?: boolean;
    value?: string;
    options?: XlideSidebarSelectOption[];
    command?: XlideSidebarCommand;
    children?: XlideSidebarNode[];
}

interface XlideSidebarModelInput {
    workbookChoices?: readonly XlideSidebarWorkbookChoice[];
    activeWorkbook?: XlideSidebarActiveWorkbook;
    setupStatus?: XlideSidebarSetupStatus;
}

function buildXlideSidebarModel(input: XlideSidebarModelInput): XlideSidebarNode[] {
    const workbookArg = input.activeWorkbook ? workbookCommandArg(input.activeWorkbook) : undefined;
    const setupStatus = input.setupStatus ?? defaultSetupStatus();
    const setupComplete = isXlideSetupComplete(setupStatus);
    const setupSections = [
        welcomeSection(setupComplete),
        section('setup', 'Setup', [
            statusNode(
                'setup.pythonExecutable',
                'Python Executable',
                setupStatus.pythonExecutable.description,
                setupStatus.pythonExecutable.status,
                setupStatus.pythonExecutable.tooltip,
                undefined,
                pythonExecutableSetupCommand(setupStatus.pythonExecutable),
                setupStatus.pythonExecutable.status === 'pass',
            ),
            statusNode(
                'setup.pythonLibraries',
                'Required Python Libraries',
                setupStatus.pythonLibraries.description,
                setupStatus.pythonLibraries.status,
                setupStatus.pythonLibraries.tooltip,
                undefined,
                pythonLibrariesSetupCommand(setupStatus.pythonLibraries),
                setupStatus.pythonLibraries.status === 'pass',
            ),
        ]),
    ];
    if (!setupComplete) {
        return setupSections;
    }
    return [
        ...setupSections,
        section('workbookActions', 'Workbook Actions', [
            targetWorkbookNode(input.workbookChoices ?? [], input.activeWorkbook),
            workbookActionNode(
                'workbookActions.analyzeWorkbook',
                'Analyze Workbook',
                undefined,
                'xlide.analyzeWorkbook',
                'Analyze the selected target workbook.',
                'checklist',
                workbookArg,
            ),
            workbookActionNode(
                'workbookActions.exportModules',
                'Export Modules',
                undefined,
                'xlide.exportModulesToFolder',
                'Open the module export diff GUI for the selected target workbook.',
                'export',
                workbookArg,
            ),
            workbookActionNode(
                'workbookActions.importModules',
                'Import Modules',
                undefined,
                'xlide.importModulesFromFolder',
                'Open the module import diff GUI for the selected target workbook.',
                'import',
                workbookArg,
            ),
            workbookActionNode(
                'workbookActions.openWorkbook',
                'Open Workbook In Excel',
                undefined,
                'xlide.openWorkbook',
                'Open the selected target workbook in Excel.',
                'file-excel',
                workbookArg,
            ),
            workbookActionNode(
                'workbookActions.openWorkbookReadOnly',
                'Open Workbook Read Only',
                undefined,
                'xlide.openWorkbookReadOnly',
                'Open the selected target workbook in Excel as read-only.',
                'file-excel',
                workbookArg,
            ),
            workbookActionNode(
                'workbookActions.runVbaTests',
                'Unit Tests',
                undefined,
                'xlide.runVbaTests',
                'Open the VBA unit tests GUI for the selected target workbook.',
                'beaker',
                workbookArg,
            ),
        ]),
        section('settings', 'Settings', [
            actionNode(
                'settings.openGlobal',
                'Global Settings',
                'VS Code / Machine',
                'xlide.openGlobalSettings',
                'Open XLIDE global/editor settings.',
                'settings-gear',
            ),
        ]),
        section('support', 'Support', [
            actionNode(
                'support.copyDiagnostics',
                'Copy Diagnostics',
                undefined,
                'xlide.copyDiagnostics',
                'Copy a redacted diagnostic snapshot to the clipboard.',
                'clippy',
            ),
            actionNode(
                'support.exportBundle',
                'Export Support Bundle',
                undefined,
                'xlide.exportSupportBundle',
                'Export a redacted support bundle for troubleshooting.',
                'archive',
            ),
        ]),
        section('donate', 'Donate', [
            actionNode(
                'donate.cashApp',
                'Donate',
                'Cash App: $williamesmithjcil 💵',
                'xlide.openCashAppDonateLink',
                'Support XLIDE with a Cash App donation.',
                'heart',
            ),
            actionNode(
                'donate.githubSponsors',
                'Donate',
                'GitHub Sponsors ❤️',
                'xlide.openSponsorLink',
                'Support XLIDE through GitHub Sponsors.',
                'heart',
            ),
            actionNode(
                'donate.paypal',
                'Donate',
                'PayPal 💳',
                'xlide.openPayPalDonateLink',
                'Support XLIDE with a PayPal donation.',
                'heart',
            ),
        ]),
    ];
}

function welcomeSection(setupComplete: boolean): XlideSidebarNode {
    if (!setupComplete) {
        return section('welcome', 'Welcome', [
            statusNode(
                'welcome.setupRequired',
                'Setup Required',
                'Please see Setup below to proceed.',
                'warn',
                'Complete Python executable and required library setup before XLIDE shows workbook navigation and actions.',
            ),
        ]);
    }
    return section('welcome', 'Welcome', [
        statusNode(
            'welcome.tree',
            'Workbook Tree',
            'Find workbook and module navigation in Explorer > XLIDE.',
            'unknown',
            'The XLIDE workbook tree stays in the VS Code Explorer so workbook navigation and sidebar actions remain separate.',
        ),
    ]);
}

function isXlideSetupComplete(setupStatus: XlideSidebarSetupStatus): boolean {
    return setupStatus.pythonExecutable.status === 'pass' && setupStatus.pythonLibraries.status === 'pass';
}

function pythonExecutableSetupCommand(status: XlideSidebarDependencyStatus): XlideSidebarCommand {
    if (status.status === 'pass') {
        return { command: 'xlide.downloadPython', title: 'Installed' };
    }
    if (status.action === 'downloadPython') {
        return {
            command: 'xlide.downloadPython',
            title: 'Download',
            tooltip: 'Setup has two gates: Python Executable and Required Python Libraries. Click to download Python from python.org. Ctrl+click to browse for an installed Python executable.',
            ctrlCommand: 'xlide.browsePythonPath',
            ctrlTitle: 'Browse',
            ctrlTooltip: 'Browse for an installed Python executable on this machine.',
        };
    }
    return {
        command: 'workbench.action.openSettings',
        title: 'Set Path',
        arguments: ['xlide.pythonPath'],
    };
}

function pythonLibrariesSetupCommand(status: XlideSidebarDependencyStatus): XlideSidebarCommand {
    return {
        command: 'xlide.setup',
        title: status.status === 'pass' ? 'Installed' : 'Install',
    };
}

function defaultSetupStatus(): XlideSidebarSetupStatus {
    return {
        pythonExecutable: {
            status: 'unknown',
            description: 'Checking',
            tooltip: 'XLIDE is checking the configured Python executable.',
        },
        pythonLibraries: {
            status: 'unknown',
            description: 'Checking',
            tooltip: 'XLIDE is checking required Python libraries.',
        },
    };
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
    command?: XlideSidebarCommand,
    disabled = false,
): XlideSidebarNode {
    return { id, kind: 'status', label, description, status, tooltip, icon, command, disabled };
}

function targetWorkbookNode(
    choices: readonly XlideSidebarWorkbookChoice[],
    workbook: XlideSidebarActiveWorkbook | undefined,
): XlideSidebarNode {
    if (choices.length > 0) {
        const optionValues = new Set(choices.map((choice) => normalizePathKey(choice.filePath)));
        const options = [
            ...(workbook ? [] : [{ label: 'Select Workbook...', value: '' }]),
            ...(workbook && !optionValues.has(normalizePathKey(workbook.filePath))
                ? [{
                    label: workbook.label,
                    value: workbook.filePath,
                    description: workbook.filePath,
                }]
                : []),
            ...choices.map((choice) => ({
                label: choice.label,
                value: choice.filePath,
                description: choice.description,
            })),
        ];
        return {
            id: 'project.targetWorkbook',
            kind: 'select',
            label: 'Target Workbook',
            description: workbook ? workbook.label : 'Select Workbook',
            status: workbook ? 'pass' : 'warn',
            tooltip: workbook
                ? `${workbook.filePath}\nSelected from ${selectionSourceLabel(workbook.selectionSource)}.`
                : 'Choose the workbook that sidebar actions should analyze, import/export, validate, run, or test.',
            icon: workbook ? 'file-code' : 'warning',
            value: workbook?.filePath ?? '',
            options,
        };
    }
    if (!workbook) {
        return statusNode(
            'project.targetWorkbook',
            'Target Workbook',
            'None Selected',
            'unknown',
            'Open an XLIDE VBA module, choose a workbook in the sidebar, or keep exactly one workbook in the workspace to select a workbook context.',
        );
    }
    return statusNode(
        'project.targetWorkbook',
        'Target Workbook',
        workbook.label,
        'pass',
        `${workbook.filePath}\nSelected from ${selectionSourceLabel(workbook.selectionSource)}.`,
        'file-code',
    );
}

function normalizePathKey(filePath: string): string {
    return filePath.toLowerCase();
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

function workbookActionNode(
    id: string,
    label: string,
    description: string | undefined,
    command: string,
    tooltip: string,
    icon: string,
    workbookArg: unknown | undefined,
): XlideSidebarNode {
    if (!workbookArg) {
        return {
            id,
            kind: 'action',
            label,
            description,
            tooltip: 'Select a target workbook before running this sidebar action.',
            icon,
            disabled: true,
        };
    }
    return actionNode(id, label, description, command, tooltip, icon, [workbookArg]);
}

function workbookCommandArg(workbook: XlideSidebarActiveWorkbook): { kind: 'xlsm'; label: string; filePath: string } {
    return {
        kind: 'xlsm',
        label: workbook.label,
        filePath: workbook.filePath,
    };
}

function selectionSourceLabel(source: XlideSidebarActiveWorkbook['selectionSource']): string {
    switch (source) {
        case 'activeEditor':
            return 'the active editor';
        case 'singleWorkbook':
            return 'the only workbook in the workspace';
        case 'sidebarSelection':
            return 'the sidebar workbook picker';
        default:
            return 'the current context';
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

export {
    buildXlideSidebarModel,
    isXlideSetupComplete,
    type XlideSidebarActiveWorkbook,
    type XlideSidebarCommand,
    type XlideSidebarDependencyStatus,
    type XlideSidebarModelInput,
    type XlideSidebarNode,
    type XlideSidebarNodeKind,
    type XlideSidebarSetupAction,
    type XlideSidebarSetupStatus,
    type XlideSidebarWorkbookChoice,
    type XlideSidebarStatus,
};
