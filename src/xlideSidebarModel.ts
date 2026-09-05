import { projectIdentityKey } from './projectIdentity';
import { containerAppNameForPath, isExcelContainerPath } from './macroContainerUi';

type XlideSidebarNodeKind = 'section' | 'status' | 'action' | 'select' | 'note' | 'link';
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
    /** A link node's destination; the host opens it, and only from SPONSOR_LINKS. */
    url?: string;
    /** A link node's mark: an inline icon name, or an emoji when there is no icon for it. */
    icon?: string;
}

interface XlideSponsorLink {
    label: string;
    detail: string;
    url: string;
    icon: string;
}

/**
 * Where to support the work, if it has been useful. The same three addresses
 * the VBA editor add-in offers, and the only ones the sidebar will open.
 */
const SPONSOR_LINKS: readonly XlideSponsorLink[] = [
    {
        label: 'GitHub Sponsors',
        detail: 'Recurring or one-off, through GitHub',
        icon: 'github',
        url: 'https://github.com/sponsors/WilliamSmithEdward',
    },
    {
        label: 'PayPal',
        detail: 'One-off, no account needed',
        icon: 'credit-card',
        url: 'https://www.paypal.com/donate/?business=ML855BRLNR838&no_recurring=0&item_name=VBA+has+always+treated+me+well.+It+was+how+I+first+grew+professional+as+a+programmer%2C+I%27m+happy+to+show+it+some+love+%E2%9D%A4%EF%B8%8F&currency_code=USD',
    },
    {
        // The banknote emoji: the inline icon set has nothing for cash, and the
        // add-in draws the same one.
        label: 'Cash App',
        detail: '$williamesmithjcil',
        icon: '\u{1F4B5}',
        url: 'https://cash.app/$williamesmithjcil',
    },
];

const SPONSOR_BLURB =
    'VBA has always treated me well. It is how I first grew professional as a programmer, '
    + 'and XLIDE is what I wish it had come with. If it has been useful, here is where to say so.';

const SPONSOR_THANKS = 'Nothing here is ever required. Thank you for using it either way.';

/**
 * The dialog's rows as quick pick items: a codicon or emoji mark, the label,
 * the detail beside it, and the address a selection opens.
 */
function sponsorQuickPickItems(): Array<{ label: string; description: string; url: string }> {
    return SPONSOR_LINKS.map((link) => ({
        label: `${/^[a-z0-9-]+$/.test(link.icon) ? `$(${link.icon})` : link.icon} ${link.label}`,
        description: link.detail,
        url: link.url,
    }));
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
        sponsorSection(),
    ];
}

function sponsorSection(): XlideSidebarNode {
    return section('sponsor', 'Support XLIDE', [
        noteNode('sponsor.blurb', SPONSOR_BLURB),
        ...SPONSOR_LINKS.map((link) => linkNode(link)),
        noteNode('sponsor.thanks', SPONSOR_THANKS),
    ]);
}

function noteNode(id: string, label: string): XlideSidebarNode {
    return { id, kind: 'note', label };
}

function linkNode(link: XlideSponsorLink): XlideSidebarNode {
    return {
        id: `sponsor.${link.label.toLowerCase().replace(/[^a-z0-9]+/g, '-')}`,
        kind: 'link',
        label: link.label,
        description: link.detail,
        tooltip: link.url,
        url: link.url,
        icon: link.icon,
    };
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
    SPONSOR_BLURB,
    SPONSOR_LINKS,
    SPONSOR_THANKS,
    sponsorQuickPickItems,
    type XlideSidebarActiveProject,
    type XlideSidebarCommand,
    type XlideSidebarModelInput,
    type XlideSidebarNode,
    type XlideSidebarNodeKind,
    type XlideSidebarProjectChoice,
    type XlideSidebarStatus,
};
