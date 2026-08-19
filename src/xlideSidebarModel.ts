import { workbookIdentityKey } from './workbookIdentity';
import { containerAppNameForPath, isExcelContainerPath } from './macroContainerUi';

type XlideSidebarNodeKind = 'section' | 'status' | 'action' | 'select';
type XlideSidebarStatus = 'pass' | 'warn' | 'fail' | 'unknown';

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

interface XlideSidebarNode {
    id: string;
    kind: XlideSidebarNodeKind;
    label: string;
    description?: string;
    tooltip?: string;
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
}

function buildXlideSidebarModel(input: XlideSidebarModelInput): XlideSidebarNode[] {
    const workbookArg = input.activeWorkbook ? workbookCommandArg(input.activeWorkbook) : undefined;
    return [
        welcomeSection(),
        section('workbookActions', 'File Actions', [
            targetWorkbookNode(input.workbookChoices ?? [], input.activeWorkbook),
            workbookActionNode(
                'workbookActions.analyzeWorkbook',
                'Analyze File',
                undefined,
                'xlide.analyzeWorkbook',
                'Analyze the selected target file.',
                workbookArg,
            ),
            workbookActionNode(
                'workbookActions.exportModules',
                'Export Modules',
                undefined,
                'xlide.exportModulesToFolder',
                'Open the module export diff GUI for the selected target file.',
                workbookArg,
            ),
            workbookActionNode(
                'workbookActions.importModules',
                'Import Modules',
                undefined,
                'xlide.importModulesFromFolder',
                'Open the module import diff GUI for the selected target file.',
                workbookArg,
            ),
            ...openActionNodes(input.activeWorkbook, workbookArg),
            workbookActionNode(
                'workbookActions.runVbaTests',
                'Unit Tests',
                undefined,
                'xlide.runVbaTests',
                'Open the VBA unit tests GUI for the selected target file.',
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
            ),
        ]),
        section('support', 'Support', [
            actionNode(
                'support.copyDiagnostics',
                'Copy Diagnostics',
                undefined,
                'xlide.copyDiagnostics',
                'Copy a redacted diagnostic snapshot to the clipboard.',
            ),
            actionNode(
                'support.exportBundle',
                'Export Support Bundle',
                undefined,
                'xlide.exportSupportBundle',
                'Export a redacted support bundle for troubleshooting.',
            ),
        ]),
    ];
}

function welcomeSection(): XlideSidebarNode {
    return section('welcome', 'Welcome', [
        statusNode(
            'welcome.tree',
            'File Tree',
            'Find file and module navigation in Explorer > XLIDE.',
            'unknown',
            'The XLIDE file tree stays in the VS Code Explorer so file navigation and sidebar actions remain separate.',
        ),
    ]);
}

function section(id: string, label: string, children: XlideSidebarNode[]): XlideSidebarNode {
    return { id, kind: 'section', label, children };
}

function statusNode(
    id: string,
    label: string,
    description: string,
    status: XlideSidebarStatus,
    tooltip: string,
    command?: XlideSidebarCommand,
    disabled = false,
): XlideSidebarNode {
    return { id, kind: 'status', label, description, status, tooltip, command, disabled };
}

function targetWorkbookNode(
    choices: readonly XlideSidebarWorkbookChoice[],
    workbook: XlideSidebarActiveWorkbook | undefined,
): XlideSidebarNode {
    if (choices.length > 0) {
        const optionValues = new Set(choices.map((choice) => workbookIdentityKey(choice.filePath)));
        const options = [
            ...(workbook ? [] : [{ label: 'Select File...', value: '' }]),
            ...(workbook && !optionValues.has(workbookIdentityKey(workbook.filePath))
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
            label: 'Target File',
            description: workbook ? workbook.label : 'Select File',
            status: workbook ? 'pass' : 'warn',
            tooltip: workbook
                ? `${workbook.filePath}\nSelected from ${selectionSourceLabel(workbook.selectionSource)}.`
                : 'Choose the file that sidebar actions should analyze, import/export, validate, run, or test.',
            value: workbook?.filePath ?? '',
            options,
        };
    }
    if (!workbook) {
        return statusNode(
            'project.targetWorkbook',
            'Target File',
            'None Selected',
            'unknown',
            'Open an XLIDE VBA module, choose a file in the sidebar, or keep exactly one macro-enabled file in the workspace to select a target.',
        );
    }
    return statusNode(
        'project.targetWorkbook',
        'Target File',
        workbook.label,
        'pass',
        `${workbook.filePath}\nSelected from ${selectionSourceLabel(workbook.selectionSource)}.`,
    );
}

/**
 * The open-in-application actions for the selected file. Excel files keep
 * the launcher pair (normal and read-only); every other host opens through
 * its own application via the OS association, which has no read-only mode.
 */
function openActionNodes(
    workbook: XlideSidebarActiveWorkbook | undefined,
    workbookArg: unknown | undefined,
): XlideSidebarNode[] {
    if (workbook && !isExcelContainerPath(workbook.filePath)) {
        const app = containerAppNameForPath(workbook.filePath);
        return [
            workbookActionNode(
                'workbookActions.openWorkbook',
                `Open in ${app}`,
                undefined,
                'xlide.openInOfficeApp',
                `Open the selected target file in ${app}.`,
                workbookArg,
            ),
        ];
    }
    return [
        workbookActionNode(
            'workbookActions.openWorkbook',
            'Open Workbook in Excel',
            undefined,
            'xlide.openWorkbook',
            'Open the selected target workbook in Excel.',
            workbookArg,
        ),
        workbookActionNode(
            'workbookActions.openWorkbookReadOnly',
            'Open Workbook in Excel (Read Only)',
            undefined,
            'xlide.openWorkbookReadOnly',
            'Open the selected target workbook in Excel as read-only.',
            workbookArg,
        ),
    ];
}

function actionNode(
    id: string,
    label: string,
    description: string | undefined,
    command: string,
    tooltip: string,
    args: unknown[] = [],
): XlideSidebarNode {
    return {
        id,
        kind: 'action',
        label,
        description,
        tooltip,
        command: { command, title: label, arguments: args },
    };
}

function workbookActionNode(
    id: string,
    label: string,
    description: string | undefined,
    command: string,
    tooltip: string,
    workbookArg: unknown | undefined,
): XlideSidebarNode {
    if (!workbookArg) {
        return {
            id,
            kind: 'action',
            label,
            description,
            tooltip: 'Select a target file before running this sidebar action.',
            disabled: true,
        };
    }
    return actionNode(id, label, description, command, tooltip, [workbookArg]);
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
            return 'the only macro-enabled file in the workspace';
        case 'sidebarSelection':
            return 'the sidebar file picker';
        default:
            return 'the current context';
    }
}

export {
    buildXlideSidebarModel,
    type XlideSidebarActiveWorkbook,
    type XlideSidebarCommand,
    type XlideSidebarModelInput,
    type XlideSidebarNode,
    type XlideSidebarNodeKind,
    type XlideSidebarWorkbookChoice,
    type XlideSidebarStatus,
};
