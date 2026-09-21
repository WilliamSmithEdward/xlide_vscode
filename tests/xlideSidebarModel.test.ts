import { describe, expect, it } from 'vitest';
import { buildXlideSidebarModel, type XlideSidebarNode } from '../src/xlideSidebarModel';

/** A section by its id, so a new section does not shift every lookup. */
function sectionOf(model: readonly XlideSidebarNode[], id: string): XlideSidebarNode {
    const found = model.find((section) => section.id === id);
    expect(found, id).toBeDefined();
    return found!;
}

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
            'Agentic AI',
            'Project Actions',
            'Settings',
            'Support',
        ]);
        expect(model[0].children?.map((node) => [node.label, node.description, node.kind])).toEqual([
            ['File Tree', 'Find file and module navigation in Explorer > XLIDE.', 'status'],
        ]);
        expect(sectionOf(model, 'projectActions').children?.map((node) => node.label)).toEqual([
            'Target File',
            'Analyze Project',
            'Export Modules',
            'Import Modules',
            'Open in Office Application',
            'Open in Office Application (Read Only)',
            'Unit Tests',
        ]);
        expect(sectionOf(model, 'settings').children?.map((node) => [node.label, node.description])).toEqual([
            ['Global Settings', 'VS Code / Machine'],
        ]);
        expect(sectionOf(model, 'settings').children?.[0]?.command?.command).toBe('xlide.openGlobalSettings');
        expect(sectionOf(model, 'support').children?.map((node) => node.label)).toEqual([
            'Copy Diagnostics',
            'Export Support Bundle',
        ]);
    });

    it('offers the agent instructions between Welcome and Project Actions, as a dialog rather than a command', () => {
        const agentic = sectionOf(buildXlideSidebarModel({}), 'agenticAi');

        expect(agentic.label).toBe('Agentic AI');
        expect(agentic.children).toEqual([expect.objectContaining({
            id: 'agenticAi.instructions',
            kind: 'action',
            label: 'Agent Instructions',
            dialog: 'agentInstructions',
        })]);
        expect(agentic.children?.[0]?.command).toBeUndefined();
    });

    it('ends with the sections the user works in, and asks for nothing after them', () => {
        const model = buildXlideSidebarModel({});

        expect(model[model.length - 1].id).toBe('support');
        expect(model.map((section) => section.id)).not.toContain('sponsor');
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
        const selector = sectionOf(model, 'projectActions').children?.find((node) => node.id === 'project.targetProject');

        expect(selector?.kind).toBe('select');
        expect(selector?.label).toBe('Target File');
        expect(selector?.value).toBe('C:\\work\\Book.xlsm');
        expect(selector?.options?.map((option) => [option.label, option.value])).toEqual([
            ['Book.xlsm', 'C:\\work\\Book.xlsm'],
        ]);
    });

    const openActions = (fileName: string) => {
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
        return (sectionOf(model, 'projectActions').children ?? [])
            .filter((node) => node.id.startsWith('projectActions.openInApp'))
            .map((node) => [node.label, node.command?.command]);
    };

    it('offers the same open pair for every application that can open read-only', () => {
        for (const [fileName, app] of [
            ['Book.xlsm', 'Excel'],
            ['Report.docm', 'Word'],
            ['Deck.pptm', 'PowerPoint'],
        ] as const) {
            expect(openActions(fileName), fileName).toEqual([
                [`Open in ${app}`, 'xlide.openInOfficeApp'],
                [`Open in ${app} (Read Only)`, 'xlide.openInOfficeAppReadOnly'],
            ]);
        }
    });

    it('offers no read-only open where there is none to offer', () => {
        // Access has no read-only open, and a VB6 project opens through
        // whatever the operating system has registered for it.
        expect(openActions('Data.accdb')).toEqual([['Open in Access', 'xlide.openInOfficeApp']]);
        expect(openActions('App.vbp')).toEqual([['Open in Visual Basic 6', 'xlide.openInOfficeApp']]);
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

        expect(sectionOf(model, 'projectActions').children?.map((node) => node.id)).not.toContain('projectActions.settingsJson');
        expect(sectionOf(model, 'projectActions').children?.map((node) => node.label)).not.toContain('Workbook Settings JSON');
        expect(sectionOf(model, 'settings').children?.map((node) => node.label)).not.toContain('Workbook Settings JSON');
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
            'projectActions.openInApp',
            'projectActions.openInAppReadOnly',
        ]) {
            expect(sectionOf(model, 'projectActions').children?.find((node) => node.id === id)?.command?.arguments).toEqual([{
                kind: 'project',
                label: 'Second.xlsm',
                filePath: 'C:\\work\\Second.xlsm',
            }]);
        }
        expect(sectionOf(model, 'projectActions').children?.map((node) => node.id)).not.toContain('projectActions.validateProject');
    });
});
