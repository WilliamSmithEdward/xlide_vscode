import { describe, expect, it } from 'vitest';
import { buildXlideSidebarModel, isSponsorUrl, SPONSOR_LINKS } from '../src/xlideSidebarModel';

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
            'Support XLIDE',
        ]);
        expect(model[0].children?.map((node) => [node.label, node.description, node.kind])).toEqual([
            ['File Tree', 'Find file and module navigation in Explorer > XLIDE.', 'status'],
        ]);
        expect(model[1].children?.map((node) => node.label)).toEqual([
            'Target File',
            'Analyze Project',
            'Export Modules',
            'Import Modules',
            'Open in Office Application',
            'Open in Office Application (Read Only)',
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

    it('ends with the sponsor section: a blurb, the three addresses, and the thanks line', () => {
        const model = buildXlideSidebarModel({});
        const sponsor = model[model.length - 1];

        expect(sponsor.id).toBe('sponsor');
        expect(sponsor.children?.map((node) => node.kind)).toEqual(['note', 'link', 'link', 'link', 'note']);
        expect(sponsor.children?.filter((node) => node.kind === 'link').map((node) => [node.label, node.description, node.url])).toEqual([
            ['GitHub Sponsors', 'Recurring or one-off, through GitHub', 'https://github.com/sponsors/WilliamSmithEdward'],
            ['PayPal', 'One-off, no account needed', SPONSOR_LINKS[1].url],
            ['Cash App', '$williamesmithjcil', 'https://cash.app/$williamesmithjcil'],
        ]);
        expect(sponsor.children?.[4]?.label).toBe('Nothing here is ever required. Thank you for using it either way.');
    });

    it('opens or copies only the three sponsor addresses', () => {
        for (const link of SPONSOR_LINKS) {
            expect(isSponsorUrl(link.url)).toBe(true);
        }
        expect(isSponsorUrl('https://github.com/sponsors/SomeoneElse')).toBe(false);
        expect(isSponsorUrl('https://cash.app/$williamesmithjcil/extra')).toBe(false);
        expect(isSponsorUrl(undefined)).toBe(false);
        expect(isSponsorUrl(42)).toBe(false);
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
        return (model[1].children ?? [])
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
            'projectActions.openInApp',
            'projectActions.openInAppReadOnly',
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
