import { describe, expect, it } from 'vitest';
import { buildXlideSidebarModel, isXlideSetupComplete } from '../src/xlideSidebarModel';

describe('xlideSidebarModel', () => {
    it('builds the sidebar sections in the product order with title-case labels', () => {
        const model = buildXlideSidebarModel({
            workbookChoices: [
                { label: 'BookA.xlsm', filePath: 'C:\\work\\BookA.xlsm' },
                { label: 'BookB.xlsm', filePath: 'C:\\work\\BookB.xlsm' },
            ],
            setupStatus: completeSetupStatus(),
        });

        expect(model.map((section) => section.label)).toEqual([
            'Welcome',
            'Setup',
            'Workbook Actions',
            'Settings',
            'Support',
        ]);
        expect(model[0].children?.map((node) => [node.label, node.description, node.kind])).toEqual([
            ['Workbook Tree', 'Find workbook and module navigation in Explorer > XLIDE.', 'status'],
        ]);
        expect(model[2].children?.map((node) => node.label)).toEqual([
            'Target Workbook',
            'Analyze Workbook',
            'Export Modules',
            'Import Modules',
            'Open Workbook In Excel',
            'Open Workbook Read Only',
            'Unit Tests',
        ]);
        expect(model[3].children?.map((node) => [node.label, node.description])).toEqual([
            ['Global Settings', 'VS Code / Machine'],
        ]);
        expect(model[3].children?.[0]?.command?.command).toBe('xlide.openGlobalSettings');
        expect(model[4].children?.map((node) => node.label)).toEqual([
            'Copy Diagnostics',
            'Export Support Bundle',
        ]);
    });

    it('renders setup rows with Installed buttons once they are green', () => {
        const model = buildXlideSidebarModel({
            setupStatus: completeSetupStatus(),
        });
        const setup = model[1];

        expect(setup.children?.map((node) => [node.id, node.label, node.description, node.status])).toEqual([
            ['setup.pythonExecutable', 'Python Executable', 'C:\\Python\\python.exe', 'pass'],
            ['setup.pythonLibraries', 'Required Python Libraries', 'Installed', 'pass'],
        ]);
        expect(setup.children?.map((node) => [node.command?.command, node.command?.title, node.disabled])).toEqual([
            ['xlide.downloadPython', 'Installed', true],
            ['xlide.setup', 'Installed', true],
        ]);
    });

    it('offers Download when Python is missing and Install when libraries are missing', () => {
        const model = buildXlideSidebarModel({
            setupStatus: {
                pythonExecutable: {
                    status: 'warn',
                    description: 'Not Found',
                    tooltip: 'Python was not found.',
                    action: 'downloadPython',
                },
                pythonLibraries: {
                    status: 'warn',
                    description: 'Missing',
                    tooltip: 'Required libraries are missing.',
                },
            },
        });

        expect(model.map((section) => section.label)).toEqual(['Welcome', 'Setup']);
        expect(model[0].children?.map((node) => [node.label, node.description])).toEqual([
            ['Setup Required', 'Please see Setup below to proceed.'],
        ]);
        expect(model[1].children?.map((node) => [node.label, node.status, node.command?.command, node.command?.title, node.disabled])).toEqual([
            ['Python Executable', 'warn', 'xlide.downloadPython', 'Download', false],
            ['Required Python Libraries', 'warn', 'xlide.setup', 'Install', false],
        ]);
        const pythonCommand = model[1].children?.[0]?.command;
        expect(pythonCommand).toMatchObject({
            tooltip: expect.stringContaining('Setup has two gates'),
            ctrlCommand: 'xlide.browsePythonPath',
            ctrlTitle: 'Browse',
        });
        expect(model[1].children?.[1]?.command?.ctrlCommand).toBeUndefined();
    });

    it('offers Set Path when Python is installed but not available to XLIDE', () => {
        const model = buildXlideSidebarModel({
            setupStatus: {
                pythonExecutable: {
                    status: 'warn',
                    description: 'Not On PATH',
                    tooltip: 'Set xlide.pythonPath.',
                    action: 'setPythonPath',
                },
                pythonLibraries: {
                    status: 'unknown',
                    description: 'Waiting For Python',
                    tooltip: 'Set Python first.',
                },
            },
        });

        expect(model[1].children?.map((node) => [node.label, node.status, node.command?.command, node.command?.title, node.disabled])).toEqual([
            ['Python Executable', 'warn', 'workbench.action.openSettings', 'Set Path', false],
            ['Required Python Libraries', 'unknown', 'xlide.setup', 'Install', false],
        ]);
        expect(model[1].children?.[0]?.command?.ctrlCommand).toBeUndefined();
    });

    it('always uses a selector for workspace workbook choices', () => {
        const model = buildXlideSidebarModel({
            workbookChoices: [
                { label: 'Book.xlsm', filePath: 'C:\\work\\Book.xlsm' },
            ],
            setupStatus: completeSetupStatus(),
            activeWorkbook: {
                label: 'Book.xlsm',
                filePath: 'C:\\work\\Book.xlsm',
                settingsPath: 'C:\\work\\Book.xlsm.xlide_settings.json',
                selectionSource: 'singleWorkbook',
                settingsState: 'valid',
            },
        });
        const selector = model[2].children?.find((node) => node.id === 'project.targetWorkbook');

        expect(selector?.kind).toBe('select');
        expect(selector?.label).toBe('Target Workbook');
        expect(selector?.value).toBe('C:\\work\\Book.xlsm');
        expect(selector?.options?.map((option) => [option.label, option.value])).toEqual([
            ['Book.xlsm', 'C:\\work\\Book.xlsm'],
        ]);
    });

    it('keeps Workbook Settings JSON out of the permanent sidebar actions', () => {
        const model = buildXlideSidebarModel({
            workbookChoices: [
                { label: 'Book.xlsm', filePath: 'C:\\work\\Book.xlsm' },
            ],
            setupStatus: completeSetupStatus(),
            activeWorkbook: {
                label: 'Book.xlsm',
                filePath: 'C:\\work\\Book.xlsm',
                settingsPath: 'C:\\work\\Book.xlsm.xlide_settings.json',
                selectionSource: 'sidebarSelection',
                settingsState: 'missing',
            },
        });

        expect(model[2].children?.map((node) => node.id)).not.toContain('workbookActions.settingsJson');
        expect(model[2].children?.map((node) => node.label)).not.toContain('Workbook Settings JSON');
        expect(model[3].children?.map((node) => node.label)).not.toContain('Workbook Settings JSON');
    });

    it('passes the selected workbook to every workbook-scoped action', () => {
        const model = buildXlideSidebarModel({
            workbookChoices: [
                { label: 'First.xlsm', filePath: 'C:\\work\\First.xlsm' },
                { label: 'Second.xlsm', filePath: 'C:\\work\\Second.xlsm' },
            ],
            setupStatus: completeSetupStatus(),
            activeWorkbook: {
                label: 'Second.xlsm',
                filePath: 'C:\\work\\Second.xlsm',
                settingsPath: 'C:\\work\\Second.xlsm.xlide_settings.json',
                selectionSource: 'sidebarSelection',
                settingsState: 'valid',
            },
        });

        for (const id of [
            'workbookActions.analyzeWorkbook',
            'workbookActions.runVbaTests',
            'workbookActions.importModules',
            'workbookActions.exportModules',
            'workbookActions.openWorkbook',
            'workbookActions.openWorkbookReadOnly',
        ]) {
            expect(model[2].children?.find((node) => node.id === id)?.command?.arguments).toEqual([{
                kind: 'xlsm',
                label: 'Second.xlsm',
                filePath: 'C:\\work\\Second.xlsm',
            }]);
        }
        expect(model[2].children?.map((node) => node.id)).not.toContain('workbookActions.validateWorkbook');
    });

    it('treats setup as complete only when both dependency rows are green', () => {
        expect(isXlideSetupComplete(completeSetupStatus())).toBe(true);
        expect(isXlideSetupComplete({
            ...completeSetupStatus(),
            pythonLibraries: {
                status: 'warn',
                description: 'Missing',
                tooltip: 'Libraries missing.',
            },
        })).toBe(false);
    });
});

describe('library update availability', () => {
    function updateAvailableStatus() {
        return {
            pythonExecutable: {
                status: 'pass' as const,
                description: 'C:\\Python\\python.exe',
                tooltip: 'Python is ready.',
            },
            pythonLibraries: {
                status: 'warn' as const,
                description: 'Update Available',
                tooltip: 'pyOpenVBA 3.0.1 -> 3.2.0',
                action: 'updateLibraries' as const,
            },
        };
    }

    it('shows a warn dot with an enabled Update button wired to the update command', () => {
        const model = buildXlideSidebarModel({ setupStatus: updateAvailableStatus() });
        const setup = model.find((section) => section.id === 'setup');
        const libraries = setup?.children?.find((node) => node.id === 'setup.pythonLibraries');
        expect(libraries).toMatchObject({
            status: 'warn',
            description: 'Update Available',
            disabled: false,
            command: { command: 'xlide.updatePythonLibraries', title: 'Update' },
        });
    });

    it('does not gate the sidebar: outdated-but-installed libraries keep setup complete', () => {
        expect(isXlideSetupComplete(updateAvailableStatus())).toBe(true);
        const model = buildXlideSidebarModel({ setupStatus: updateAvailableStatus() });
        expect(model.map((section) => section.label)).toContain('Workbook Actions');
    });

    it('still gates setup when libraries are genuinely missing (warn without the action)', () => {
        const status = updateAvailableStatus();
        status.pythonLibraries = {
            status: 'warn' as const,
            description: 'Missing',
            tooltip: 'Required libraries are missing.',
        } as typeof status.pythonLibraries;
        expect(isXlideSetupComplete(status)).toBe(false);
    });
});

function completeSetupStatus() {
    return {
        pythonExecutable: {
            status: 'pass' as const,
            description: 'C:\\Python\\python.exe',
            tooltip: 'Python is ready.',
        },
        pythonLibraries: {
            status: 'pass' as const,
            description: 'Installed',
            tooltip: 'Required libraries are installed.',
        },
    };
}
