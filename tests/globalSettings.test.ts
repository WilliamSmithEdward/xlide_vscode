import { describe, expect, it } from 'vitest';
import type * as vscode from 'vscode';
import {
    normalizeXlideOptionExplicitSetting,
    resolvedXlideGlobalSettingsFromConfig,
    xlideAnalysisVisibleSeveritiesFromConfig,
    xlideDocsMetadataGlobFromConfig,
    validateXlideGlobalSettingsValues,
} from '../src/globalSettings';

const validSettings = {
    pythonPath: '',
    attachToRunningExcel: true,
    'diagnostics.enabled': true,
    'diagnostics.optionExplicit': 'warning',
    'analysis.visibleSeverities': ['error', 'warning', 'information'],
    'analysis.untrackedRules': [],
    'editor.blockLayout': 'comfy',
    'docs.enabled': true,
    'docs.metadataGlob': '**/*.vbref.xml',
};

describe('globalSettings', () => {
    it('accepts the extension default settings shape', () => {
        expect(validateXlideGlobalSettingsValues(validSettings)).toEqual([]);
    });

    it('reports malformed global analysis settings explicitly', () => {
        expect(validateXlideGlobalSettingsValues({
            ...validSettings,
            'analysis.visibleSeverities': ['error', 'hint'],
            'analysis.untrackedRules': ['option-explicit-missing', ''],
        })).toEqual([
            {
                key: 'xlide.analysis.visibleSeverities',
                message: 'Expected "xlide.analysis.visibleSeverities" entries to be one of: error, warning, information.',
                severity: 'error',
            },
            {
                key: 'xlide.analysis.untrackedRules',
                message: 'Expected "xlide.analysis.untrackedRules" entries to be non-empty strings.',
                severity: 'error',
            },
        ]);
    });

    it('reports malformed non-analysis settings through the same contract', () => {
        expect(validateXlideGlobalSettingsValues({
            ...validSettings,
            attachToRunningExcel: 'yes',
            'diagnostics.optionExplicit': 'loud',
            'editor.blockLayout': 'spacious',
        }).map((problem) => problem.key)).toEqual([
            'xlide.attachToRunningExcel',
            'xlide.diagnostics.optionExplicit',
            'xlide.editor.blockLayout',
        ]);
    });

    it('does not expose hint as a diagnostic setting severity', () => {
        expect(validateXlideGlobalSettingsValues({
            ...validSettings,
            'diagnostics.optionExplicit': 'hint',
        }).map((problem) => problem.key)).toEqual(['xlide.diagnostics.optionExplicit']);
    });

    it('normalizes invalid Option Explicit severity to the default runtime value', () => {
        expect(normalizeXlideOptionExplicitSetting('off')).toBe('off');
        expect(normalizeXlideOptionExplicitSetting('error')).toBe('error');
        expect(normalizeXlideOptionExplicitSetting('loud')).toBe('warning');
        expect(normalizeXlideOptionExplicitSetting(undefined)).toBe('warning');
    });

    it('resolves normalized global settings with explicit provenance', () => {
        const config = fakeConfig({
            'analysis.visibleSeverities': ['warning', 'hint', 'error'],
            'docs.metadataGlob': ' custom/*.vbref.xml ',
        }, new Set(['analysis.visibleSeverities']));

        expect(xlideAnalysisVisibleSeveritiesFromConfig(config)).toEqual({
            key: 'xlide.analysis.visibleSeverities',
            value: ['warning', 'error'],
            source: 'machine',
        });
        expect(xlideDocsMetadataGlobFromConfig(config)).toEqual({
            key: 'xlide.docs.metadataGlob',
            value: 'custom/*.vbref.xml',
            source: 'default',
        });
    });

    it('returns support-bundle-ready settings in the contributed key order', () => {
        const settings = resolvedXlideGlobalSettingsFromConfig(fakeConfig({
            pythonPath: 'C:\\Python\\python.exe',
        }, new Set(['pythonPath'])));

        expect(settings.map((setting) => setting.key)).toEqual([
            'xlide.analysis.untrackedRules',
            'xlide.analysis.visibleSeverities',
            'xlide.attachToRunningExcel',
            'xlide.diagnostics.enabled',
            'xlide.diagnostics.optionExplicit',
            'xlide.docs.enabled',
            'xlide.docs.metadataGlob',
            'xlide.editor.blockLayout',
            'xlide.pythonPath',
        ]);
        expect(settings.find((setting) => setting.key === 'xlide.pythonPath')).toMatchObject({
            value: 'C:\\Python\\python.exe',
            source: 'machine',
        });
    });
});

function fakeConfig(
    values: Record<string, unknown>,
    machineKeys = new Set<string>(),
): vscode.WorkspaceConfiguration {
    return {
        get: (key: string, fallback?: unknown) => key in values ? values[key] : fallback,
        inspect: (key: string) => machineKeys.has(key) ? { globalValue: values[key] } : {},
    } as unknown as vscode.WorkspaceConfiguration;
}
