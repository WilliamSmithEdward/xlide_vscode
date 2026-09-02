import { describe, expect, it } from 'vitest';
import { buildXlideSidebarModel } from '../src/xlideSidebarModel';

describe('xlideSidebarModel', () => {
    it('builds the sidebar sections in the product order with title-case labels', () => {
        const model = buildXlideSidebarModel({
            projectChoices: [
                { label: 'BookA.xlsm', filePath: 'C:\\work\\BookA.xlsm' },
                { label: 'BookB.xlsm', filePath: 'C:\\work\\BookB.xlsm' },
            ],
        });

        expect(model.map((section) => section.label)).toEqual([
            'Welcome',
            'Project Actions',
            'Settings',
            'Support',
        ]);
        expect(model[0].children?.map((node) => [node.label, node.description, node.kind])).toEqual([
            ['File Tree', 'Find file and module navigation in Explorer > XLIDE.', 'status'],
        ]);
        expect(model[1].children?.map((node) => node.label)).toEqual([
            'Target File',
            'Analyze Project',
            'Export Modules',
            'Import Modules',
            'Open Workbook in Excel',
            'Open Workbook in Excel (Read Only)',
            'Unit Tests',
        ]);
        expect(model[2].children?.map((node) => [node.label, node.description])).toEqual([
            ['Global Settings', 'VS Code / Machine'],
        ]);
        expect(model[2].children?.[0]?.command?.command).toBe('xlide.openGlobalSettings');
        expect(model[3].children?.map((node) => node.label)).toEqual([
            'Copy Diagnostics',
            'Export Support Bundle',
        ]);
    });

    it('never gates the sidebar behind a setup section', () => {
        // The workbook engine runs in-process, so there is nothing to install
        // or probe: the file actions are available from the first render.
        const model = buildXlideSidebarModel({});

        expect(model.map((section) => section.id)).not.toContain('setup');
        expect(model.map((section) => section.label)).toContain('Project Actions');
        expect(model[0].children?.map((node) => node.label)).toEqual(['File Tree']);
    });

    it('always uses a selector for workspace file choices', () => {
        const model = buildXlideSidebarModel({
            projectChoices: [
                { label: 'Book.xlsm', filePath: 'C:\\work\\Book.xlsm' },
            ],
            activeProject: {
                label: 'Book.xlsm',
                filePath: 'C:\\work\\Book.xlsm',
                settingsPath: 'C:\\work\\Book.xlsm.xlide_settings.json',
                selectionSource: 'singleProject',
                settingsState: 'valid',
            },
        });
        const selector = model[1].children?.find((node) => node.id === 'project.targetProject');

        expect(selector?.kind).toBe('select');
        expect(selector?.label).toBe('Target File');
        expect(selector?.value).toBe('C:\\work\\Book.xlsm');
        expect(selector?.options?.map((option) => [option.label, option.value])).toEqual([
            ['Book.xlsm', 'C:\\work\\Book.xlsm'],
        ]);
    });

    it('offers the Excel launcher pair only for Excel files', () => {
        const model = buildXlideSidebarModel({
            projectChoices: [
                { label: 'Report.docm', filePath: 'C:\\work\\Report.docm' },
            ],
            activeProject: {
                label: 'Report.docm',
                filePath: 'C:\\work\\Report.docm',
                settingsPath: 'C:\\work\\Report.docm.xlide_settings.json',
                selectionSource: 'sidebarSelection',
                settingsState: 'valid',
            },
        });

        const labels = model[1].children?.map((node) => node.label);
        expect(labels).toContain('Open in Word');
        expect(labels).not.toContain('Open Workbook in Excel');
        expect(labels).not.toContain('Open Workbook in Excel (Read Only)');
        const open = model[1].children?.find((node) => node.label === 'Open in Word');
        expect(open?.command?.command).toBe('xlide.openInOfficeApp');
    });

    it('names the owning application for PowerPoint and Access files', () => {
        for (const [fileName, app] of [
            ['Deck.pptm', 'PowerPoint'],
            ['Data.accdb', 'Access'],
        ] as const) {
            const model = buildXlideSidebarModel({
                projectChoices: [{ label: fileName, filePath: `C:\\work\\${fileName}` }],
                activeProject: {
                    label: fileName,
                    filePath: `C:\\work\\${fileName}`,
                    settingsPath: `C:\\work\\${fileName}.xlide_settings.json`,
                    selectionSource: 'sidebarSelection',
                    settingsState: 'valid',
                },
            });
            expect(model[1].children?.map((node) => node.label)).toContain(`Open in ${app}`);
        }
    });

    it('keeps Workbook Settings JSON out of the permanent sidebar actions', () => {
        const model = buildXlideSidebarModel({
            projectChoices: [
                { label: 'Book.xlsm', filePath: 'C:\\work\\Book.xlsm' },
            ],
            activeProject: {
                label: 'Book.xlsm',
                filePath: 'C:\\work\\Book.xlsm',
                settingsPath: 'C:\\work\\Book.xlsm.xlide_settings.json',
                selectionSource: 'sidebarSelection',
                settingsState: 'missing',
            },
        });

        expect(model[1].children?.map((node) => node.id)).not.toContain('projectActions.settingsJson');
        expect(model[1].children?.map((node) => node.label)).not.toContain('Workbook Settings JSON');
        expect(model[2].children?.map((node) => node.label)).not.toContain('Workbook Settings JSON');
    });

    it('passes the selected file to every file-scoped action', () => {
        const model = buildXlideSidebarModel({
            projectChoices: [
                { label: 'First.xlsm', filePath: 'C:\\work\\First.xlsm' },
                { label: 'Second.xlsm', filePath: 'C:\\work\\Second.xlsm' },
            ],
            activeProject: {
                label: 'Second.xlsm',
                filePath: 'C:\\work\\Second.xlsm',
                settingsPath: 'C:\\work\\Second.xlsm.xlide_settings.json',
                selectionSource: 'sidebarSelection',
                settingsState: 'valid',
            },
        });

        for (const id of [
            'projectActions.analyzeProject',
            'projectActions.runVbaTests',
            'projectActions.importModules',
            'projectActions.exportModules',
            'projectActions.openWorkbook',
            'projectActions.openWorkbookReadOnly',
        ]) {
            expect(model[1].children?.find((node) => node.id === id)?.command?.arguments).toEqual([{
                kind: 'project',
                label: 'Second.xlsm',
                filePath: 'C:\\work\\Second.xlsm',
            }]);
        }
        expect(model[1].children?.map((node) => node.id)).not.toContain('projectActions.validateProject');
    });
});
