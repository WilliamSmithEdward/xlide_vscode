import { projectIdentityKey } from './projectIdentity';
import { containerAppNameForPath, isExcelContainerPath } from './macroContainerUi';

type XlideSidebarNodeKind = 'section' | 'status' | 'action' | 'select';
type XlideSidebarStatus = 'pass' | 'warn' | 'fail' | 'unknown';

interface XlideSidebarActiveProject {
    label: string;
    filePath: string;
    settingsPath: string;
    selectionSource: 'activeEditor' | 'singleProject' | 'sidebarSelection';
    settingsState: 'missing' | 'valid' | 'invalid';
    settingsMessage?: string;
}

interface XlideSidebarProjectChoice {
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
    projectChoices?: readonly XlideSidebarProjectChoice[];
    activeProject?: XlideSidebarActiveProject;
}

function buildXlideSidebarModel(input: XlideSidebarModelInput): XlideSidebarNode[] {
    const projectArg = input.activeProject ? projectCommandArg(input.activeProject) : undefined;
    return [
        welcomeSection(),
        section('projectActions', 'Project Actions', [
            targetProjectNode(input.projectChoices ?? [], input.activeProject),
            projectActionNode(
                'projectActions.analyzeProject',
                'Analyze Project',
                undefined,
                'xlide.analyzeProject',
                'Analyze the selected target file.',
                projectArg,
            ),
            projectActionNode(
                'projectActions.exportModules',
                'Export Modules',
                undefined,
                'xlide.exportModulesToFolder',
                'Open the module export diff GUI for the selected target file.',
                projectArg,
            ),
            projectActionNode(
                'projectActions.importModules',
                'Import Modules',
                undefined,
                'xlide.importModulesFromFolder',
                'Open the module import diff GUI for the selected target file.',
                projectArg,
            ),
            ...openActionNodes(input.activeProject, projectArg),
            projectActionNode(
                'projectActions.runVbaTests',
                'Unit Tests',
                undefined,
                'xlide.runVbaTests',
                'Open the VBA unit tests GUI for the selected target file.',
                projectArg,
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

function targetProjectNode(
    choices: readonly XlideSidebarProjectChoice[],
    project: XlideSidebarActiveProject | undefined,
): XlideSidebarNode {
    if (choices.length > 0) {
        const optionValues = new Set(choices.map((choice) => projectIdentityKey(choice.filePath)));
        const options = [
            ...(project ? [] : [{ label: 'Select Project...', value: '' }]),
            ...(project && !optionValues.has(projectIdentityKey(project.filePath))
                ? [{
                    label: project.label,
                    value: project.filePath,
                    description: project.filePath,
                }]
                : []),
            ...choices.map((choice) => ({
                label: choice.label,
                value: choice.filePath,
                description: choice.description,
            })),
        ];
        return {
            id: 'project.targetProject',
            kind: 'select',
            label: 'Target File',
            description: project ? project.label : 'Select Project',
            status: project ? 'pass' : 'warn',
            tooltip: project
                ? `${project.filePath}\nSelected from ${selectionSourceLabel(project.selectionSource)}.`
                : 'Choose the file that sidebar actions should analyze, import/export, validate, run, or test.',
            value: project?.filePath ?? '',
            options,
        };
    }
    if (!project) {
        return statusNode(
            'project.targetProject',
            'Target File',
            'None Selected',
            'unknown',
            'Open an XLIDE VBA module, choose a file in the sidebar, or keep exactly one macro-enabled file in the workspace to select a target.',
        );
    }
    return statusNode(
        'project.targetProject',
        'Target File',
        project.label,
        'pass',
        `${project.filePath}\nSelected from ${selectionSourceLabel(project.selectionSource)}.`,
    );
}

/**
 * The open-in-application actions for the selected file. Excel files keep
 * the launcher pair (normal and read-only); every other host opens through
 * its own application via the OS association, which has no read-only mode.
 */
function openActionNodes(
    project: XlideSidebarActiveProject | undefined,
    projectArg: unknown | undefined,
): XlideSidebarNode[] {
    if (project && !isExcelContainerPath(project.filePath)) {
        const app = containerAppNameForPath(project.filePath);
        return [
            projectActionNode(
                'projectActions.openWorkbook',
                `Open in ${app}`,
                undefined,
                'xlide.openInOfficeApp',
                `Open the selected target file in ${app}.`,
                projectArg,
            ),
        ];
    }
    return [
        projectActionNode(
            'projectActions.openWorkbook',
            'Open Workbook in Excel',
            undefined,
            'xlide.openWorkbook',
            'Open the selected target workbook in Excel.',
            projectArg,
        ),
        projectActionNode(
            'projectActions.openWorkbookReadOnly',
            'Open Workbook in Excel (Read Only)',
            undefined,
            'xlide.openWorkbookReadOnly',
            'Open the selected target workbook in Excel as read-only.',
            projectArg,
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

function projectActionNode(
    id: string,
    label: string,
    description: string | undefined,
    command: string,
    tooltip: string,
    projectArg: unknown | undefined,
): XlideSidebarNode {
    if (!projectArg) {
        return {
            id,
            kind: 'action',
            label,
            description,
            tooltip: 'Select a target file before running this sidebar action.',
            disabled: true,
        };
    }
    return actionNode(id, label, description, command, tooltip, [projectArg]);
}

function projectCommandArg(project: XlideSidebarActiveProject): { kind: 'project'; label: string; filePath: string } {
    return {
        kind: 'project',
        label: project.label,
        filePath: project.filePath,
    };
}

function selectionSourceLabel(source: XlideSidebarActiveProject['selectionSource']): string {
    switch (source) {
        case 'activeEditor':
            return 'the active editor';
        case 'singleProject':
            return 'the only macro-enabled file in the workspace';
        case 'sidebarSelection':
            return 'the sidebar file picker';
        default:
            return 'the current context';
    }
}

export {
    buildXlideSidebarModel,
    type XlideSidebarActiveProject,
    type XlideSidebarCommand,
    type XlideSidebarModelInput,
    type XlideSidebarNode,
    type XlideSidebarNodeKind,
    type XlideSidebarProjectChoice,
    type XlideSidebarStatus,
};
