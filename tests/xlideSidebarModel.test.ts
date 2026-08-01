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
            'Workbook Actions',
            'Settings',
            'Support',
        ]);
        expect(model[0].children?.map((node) => [node.label, node.description, node.kind])).toEqual([
            ['Workbook Tree', 'Find workbook and module navigation in Explorer > XLIDE.', 'status'],
        ]);
        expect(model[1].children?.map((node) => node.label)).toEqual([
            'Target Workbook',
            'Analyze Workbook',
            'Export Modules',
            'Import Modules',
            'Open Workbook In Excel',
            'Open Workbook Read Only',
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
        // or probe: the workbook actions are available from the first render.
        const model = buildXlideSidebarModel({});

        expect(model.map((section) => section.id)).not.toContain('setup');
        expect(model.map((section) => section.label)).toContain('Workbook Actions');
        expect(model[0].children?.map((node) => node.label)).toEqual(['Workbook Tree']);
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
        const selector = model[1].children?.find((node) => node.id === 'project.targetWorkbook');

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

        expect(model[1].children?.map((node) => node.id)).not.toContain('workbookActions.settingsJson');
        expect(model[1].children?.map((node) => node.label)).not.toContain('Workbook Settings JSON');
        expect(model[2].children?.map((node) => node.label)).not.toContain('Workbook Settings JSON');
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
            'workbookActions.runVbaTests',
            'workbookActions.importModules',
            'workbookActions.exportModules',
            'workbookActions.openWorkbook',
            'workbookActions.openWorkbookReadOnly',
        ]) {
            expect(model[1].children?.find((node) => node.id === id)?.command?.arguments).toEqual([{
                kind: 'xlsm',
                label: 'Second.xlsm',
                filePath: 'C:\\work\\Second.xlsm',
            }]);
        }
        expect(model[1].children?.map((node) => node.id)).not.toContain('workbookActions.validateWorkbook');
    });
});
