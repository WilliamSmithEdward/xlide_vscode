import { describe, expect, it } from 'vitest';
import type { ResolvedXlideGlobalSetting } from '../src/globalSettings';
import { buildXlideSidebarModel, settingDescription, sourceLabel } from '../src/xlideSidebarModel';

describe('xlideSidebarModel', () => {
    it('builds native sidebar sections from the shared global settings resolver output', () => {
        const model = buildXlideSidebarModel({
            globalSettings: settings(),
            hasWorkspace: true,
            workbookCount: 2,
        });

        expect(model.map((section) => section.label)).toEqual([
            'Project',
            'Actions',
            'Configuration',
            'Support',
        ]);
        expect(model[0].children?.map((node) => [node.label, node.description])).toEqual([
            ['Workspace', 'Open'],
            ['Workbook discovery', '2 workbooks'],
            ['Reveal workbook tree', 'Explorer'],
            ['Refresh workbooks', undefined],
        ]);
        expect(model[1].children?.map((node) => node.command?.command)).toContain('xlide.analyzeWorkbook');
    });

    it('renders setting provenance without adding another settings source', () => {
        const configuration = buildXlideSidebarModel({
            globalSettings: settings(),
            hasWorkspace: true,
            workbookCount: 1,
        }).find((section) => section.id === 'configuration');

        expect(configuration?.children?.map((node) => [node.label, node.description])).toContainEqual([
            'Live diagnostics',
            'Enabled (VS Code)',
        ]);
        expect(configuration?.children?.map((node) => [node.label, node.description])).toContainEqual([
            'Python executable',
            'From PATH (Default)',
        ]);
        expect(configuration?.children?.map((node) => node.command?.command)).toContain('workbench.action.openSettings');
    });

    it('shows deterministic project status when no workspace or workbook is available', () => {
        const project = buildXlideSidebarModel({
            globalSettings: settings(),
            hasWorkspace: false,
            workbookCount: 0,
        })[0];

        expect(project.children?.map((node) => [node.label, node.description, node.status])).toEqual([
            ['Workspace', 'No folder', 'warn'],
            ['Workbook discovery', 'None found', 'unknown'],
            ['Reveal workbook tree', 'Explorer', undefined],
            ['Refresh workbooks', undefined, undefined],
        ]);
    });

    it('formats settings and sources consistently', () => {
        expect(settingDescription(true)).toBe('Enabled');
        expect(settingDescription(false)).toBe('Disabled');
        expect(settingDescription('')).toBe('From PATH');
        expect(settingDescription(['warning', 'error'])).toBe('warning, error');
        expect(settingDescription({ a: 'off', b: 'warning' })).toBe('2 overrides');
        expect(sourceLabel('machine')).toBe('VS Code');
        expect(sourceLabel('default')).toBe('Default');
    });
});

function settings(): ResolvedXlideGlobalSetting<unknown>[] {
    return [
        { key: 'xlide.analysis.ruleSeverityOverrides', value: {}, source: 'default' },
        { key: 'xlide.analysis.untrackedRules', value: [], source: 'default' },
        { key: 'xlide.analysis.visibleSeverities', value: ['error', 'warning', 'information'], source: 'default' },
        { key: 'xlide.attachToRunningExcel', value: true, source: 'default' },
        { key: 'xlide.diagnostics.enabled', value: true, source: 'machine' },
        { key: 'xlide.docs.enabled', value: true, source: 'default' },
        { key: 'xlide.docs.metadataGlob', value: '**/*.vbref.xml', source: 'default' },
        { key: 'xlide.editor.blockLayout', value: 'comfy', source: 'default' },
        { key: 'xlide.pythonPath', value: '', source: 'default' },
    ];
}
