import { describe, expect, it } from 'vitest';
import { buildXlideSidebarModel } from '../src/xlideSidebarModel';

describe('xlideSidebarModel', () => {
    it('builds the sidebar sections in the product order with title-case labels', () => {
        const model = buildXlideSidebarModel({
            workbookChoices: [
                { label: 'BookA.xlsm', filePath: 'C:\\work\\BookA.xlsm' },
                { label: 'BookB.xlsm', filePath: 'C:\\work\\BookB.xlsm' },
            ],
        });

        expect(model.map((section) => section.label)).toEqual([
            'Welcome',
            'Setup',
            'Workbook Actions',
            'Settings',
            'Support',
            'Donate',
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
        ]);
        expect(model[3].children?.map((node) => [node.label, node.description])).toEqual([
            ['Global Settings', 'VS Code / Machine'],
        ]);
        expect(model[3].children?.[0]?.command?.command).toBe('xlide.openGlobalSettings');
        expect(model[4].children?.map((node) => node.label)).toEqual([
            'Copy Diagnostics',
            'Export Support Bundle',
        ]);
        expect(model[5].children?.map((node) => [node.label, node.description])).toEqual([
            ['Donate', 'GitHub Sponsors ❤️'],
            ['Donate', 'PayPal 💳'],
            ['Donate', 'Cash App: $williamesmithjcil 💵'],
        ]);
        expect(model[5].children?.[0]?.command?.command).toBe('xlide.openSponsorLink');
        expect(model[5].children?.[1]?.command?.command).toBe('xlide.openPayPalDonateLink');
        expect(model[5].children?.[2]?.command?.command).toBe('xlide.openCashAppDonateLink');
    });

    it('renders two setup rows with disabled action buttons once they are green', () => {
        const model = buildXlideSidebarModel({
            setupStatus: {
                pythonExecutable: {
                    status: 'pass',
                    description: 'C:\\Python\\python.exe',
                    tooltip: 'Python is ready.',
                },
                pythonLibraries: {
                    status: 'pass',
                    description: 'Installed',
                    tooltip: 'Required libraries are installed.',
                },
            },
        });
        const setup = model[1];

        expect(setup.children?.map((node) => [node.id, node.label, node.description, node.status])).toEqual([
            ['setup.pythonExecutable', 'Python Executable', 'C:\\Python\\python.exe', 'pass'],
            ['setup.pythonLibraries', 'Required Python Libraries', 'Installed', 'pass'],
        ]);
        expect(setup.children?.map((node) => [node.command?.command, node.command?.title, node.disabled])).toEqual([
            ['workbench.action.openSettings', 'Set Path', true],
            ['xlide.setup', 'Install', true],
        ]);
    });

    it('enables setup action buttons when Python health needs attention', () => {
        const model = buildXlideSidebarModel({
            setupStatus: {
                pythonExecutable: {
                    status: 'warn',
                    description: 'Not Found',
                    tooltip: 'Python was not found.',
                },
                pythonLibraries: {
                    status: 'warn',
                    description: 'Missing',
                    tooltip: 'Required libraries are missing.',
                },
            },
        });

        expect(model[1].children?.map((node) => [node.label, node.status, node.disabled])).toEqual([
            ['Python Executable', 'warn', false],
            ['Required Python Libraries', 'warn', false],
        ]);
    });

    it('always uses a selector for workspace workbook choices', () => {
        const model = buildXlideSidebarModel({
            workbookChoices: [
                { label: 'Book.xlsm', filePath: 'C:\\work\\Book.xlsm' },
            ],
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
});
