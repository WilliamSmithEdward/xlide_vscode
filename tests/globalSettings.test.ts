import { describe, expect, it } from 'vitest';
import type * as vscode from 'vscode';
import {
    clearXlideGlobalAnalysisRuleSeverityOverride,
    resetXlideGlobalSettingValue,
    resolvedXlideGlobalSettingsFromConfig,
    setXlideGlobalAnalysisRuleSeverityOverride,
    setXlideGlobalAnalysisRuleTracked,
    setXlideGlobalSettingValue,
    xlideAnalysisRuleSeveritiesFromConfig,
    xlideAnalysisVisibleSeveritiesFromConfig,
    xlideDocsMetadataGlobFromConfig,
    validateXlideGlobalSettingsValues,
} from '../src/globalSettings';

const validSettings = {
    pythonPath: '',
    attachToRunningExcel: true,
    'excelIntegration.coordinationMode': 'block',
    'excelIntegration.trackOpenedWorkbooks': true,
    'excelIntegration.reopenAfterClose': true,
    'excelIntegration.reopenMode': 'readOnly',
    'excelIntegration.reopenReadOnlyAfterSave': false,
    'diagnostics.enabled': true,
    'analysis.ruleSeverityOverrides': {},
    'analysis.visibleSeverities': ['error', 'warning', 'information'],
    'analysis.untrackedRules': [],
    'editor.blockLayout': 'comfy',
    'editor.continueCommentOnNewline': true,
    'editor.mirrorCommentSpacing': true,
    'docs.enabled': true,
    'docs.metadataGlob': '**/*.vbref.xml',
    'performance.trace': false,
};

describe('globalSettings', () => {
    it('accepts the extension default settings shape', () => {
        expect(validateXlideGlobalSettingsValues(validSettings)).toEqual([]);
    });

    it('reports malformed global analysis settings explicitly', () => {
        expect(validateXlideGlobalSettingsValues({
            ...validSettings,
            'analysis.visibleSeverities': ['error', 'hint'],
            'analysis.untrackedRules': ['option-explicit-missing', 'not-a-rule'],
            'analysis.ruleSeverityOverrides': {
                'option-explicit-missing': 'error',
                'not-a-rule': 'warning',
            },
        })).toEqual([
            {
                key: 'xlide.analysis.ruleSeverityOverrides',
                message: 'Expected "xlide.analysis.ruleSeverityOverrides.option-explicit-missing" to be one of: off.',
                severity: 'error',
            },
            {
                key: 'xlide.analysis.ruleSeverityOverrides',
                message: 'Expected "xlide.analysis.ruleSeverityOverrides.not-a-rule" to target a known analysis rule that permits severity overrides.',
                severity: 'error',
            },
            {
                key: 'xlide.analysis.untrackedRules',
                message: 'Expected "xlide.analysis.untrackedRules" entries to be known analysis rule codes.',
                severity: 'error',
            },
            {
                key: 'xlide.analysis.visibleSeverities',
                message: 'Expected "xlide.analysis.visibleSeverities" entries to be one of: error, warning, information.',
                severity: 'error',
            },
        ]);
    });

    it('reports malformed non-analysis settings through the same contract', () => {
        expect(validateXlideGlobalSettingsValues({
            ...validSettings,
            attachToRunningExcel: 'yes',
            'editor.blockLayout': 'spacious',
        }).map((problem) => problem.key)).toEqual([
            'xlide.attachToRunningExcel',
            'xlide.editor.blockLayout',
        ]);
    });

    it('does not expose hint as an analysis setting severity', () => {
        expect(validateXlideGlobalSettingsValues({
            ...validSettings,
            'analysis.ruleSeverityOverrides': { 'option-explicit-missing': 'hint' },
        }).map((problem) => problem.key)).toEqual(['xlide.analysis.ruleSeverityOverrides']);
    });

    it('resolves normalized global settings with explicit provenance', () => {
        const config = fakeConfig({
            'analysis.visibleSeverities': ['warning', 'hint', 'error'],
            'analysis.ruleSeverityOverrides': {
                ' Unknown-Call ': 'warning',
                'option-explicit-missing': 'error',
            },
            'docs.metadataGlob': ' custom/*.vbref.xml ',
        }, new Set(['analysis.visibleSeverities', 'analysis.ruleSeverityOverrides']));

        expect(xlideAnalysisVisibleSeveritiesFromConfig(config)).toEqual({
            key: 'xlide.analysis.visibleSeverities',
            value: ['warning', 'error'],
            source: 'machine',
        });
        expect(xlideAnalysisRuleSeveritiesFromConfig(config)).toEqual({
            key: 'xlide.analysis.ruleSeverityOverrides',
            value: { 'unknown-call': 'warning' },
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
            'xlide.analysis.ruleSeverityOverrides',
            'xlide.analysis.untrackedRules',
            'xlide.analysis.visibleSeverities',
            'xlide.attachToRunningExcel',
            'xlide.diagnostics.enabled',
            'xlide.docs.enabled',
            'xlide.docs.metadataGlob',
            'xlide.editor.blockLayout',
            'xlide.editor.continueCommentOnNewline',
            'xlide.editor.mirrorCommentSpacing',
            'xlide.excelIntegration.coordinationMode',
            'xlide.excelIntegration.reopenAfterClose',
            'xlide.excelIntegration.reopenMode',
            'xlide.excelIntegration.reopenReadOnlyAfterSave',
            'xlide.excelIntegration.trackOpenedWorkbooks',
            'xlide.performance.trace',
            'xlide.pythonPath',
        ]);
        expect(settings.find((setting) => setting.key === 'xlide.pythonPath')).toMatchObject({
            value: 'C:\\Python\\python.exe',
            source: 'machine',
        });
    });

    it('updates global analysis rule tracking through the machine settings target', async () => {
        const updates: Array<{ key: string; value: unknown; target: unknown }> = [];
        const config = fakeConfig({
            'analysis.untrackedRules': ['argument-count'],
        }, new Set(['analysis.untrackedRules']), updates);

        await expect(setXlideGlobalAnalysisRuleTracked(config, ' Option-Explicit-Missing ', false))
            .resolves.toEqual({
                code: 'option-explicit-missing',
                tracked: false,
                changed: true,
                untrackedRules: ['argument-count', 'option-explicit-missing'],
            });
        expect(updates).toEqual([{
            key: 'analysis.untrackedRules',
            value: ['argument-count', 'option-explicit-missing'],
            target: true,
        }]);
    });

    it('skips global analysis rule tracking writes when the normalized list is unchanged', async () => {
        const updates: Array<{ key: string; value: unknown; target: unknown }> = [];
        const config = fakeConfig({
            'analysis.untrackedRules': ['option-explicit-missing'],
        }, new Set(['analysis.untrackedRules']), updates);

        await expect(setXlideGlobalAnalysisRuleTracked(config, 'option-explicit-missing', false))
            .resolves.toMatchObject({
                code: 'option-explicit-missing',
                tracked: false,
                changed: false,
                untrackedRules: ['option-explicit-missing'],
            });
        expect(updates).toEqual([]);
    });

    it('writes global setting values through the machine settings target', async () => {
        const updates: Array<{ key: string; value: unknown; target: unknown }> = [];
        const config = fakeConfig({}, new Set(), updates);

        await expect(setXlideGlobalSettingValue(config, 'analysis.untrackedRules', [
            ' Option-Explicit-Missing ',
            'not-a-rule',
            'argument-count',
        ])).resolves.toEqual({
            key: 'xlide.analysis.untrackedRules',
            value: ['argument-count', 'option-explicit-missing'],
            changed: true,
        });
        expect(updates).toEqual([{
            key: 'analysis.untrackedRules',
            value: ['argument-count', 'option-explicit-missing'],
            target: true,
        }]);
    });

    it('updates and clears guarded global rule severity overrides', async () => {
        const updates: Array<{ key: string; value: unknown; target: unknown }> = [];
        const config = fakeConfig({
            'analysis.ruleSeverityOverrides': { 'unknown-call': 'warning' },
        }, new Set(['analysis.ruleSeverityOverrides']), updates);

        await expect(setXlideGlobalAnalysisRuleSeverityOverride(
            config,
            'option-explicit-missing',
            'off',
        )).resolves.toEqual({
            key: 'xlide.analysis.ruleSeverityOverrides',
            value: {
                'option-explicit-missing': 'off',
                'unknown-call': 'warning',
            },
            changed: true,
        });
        await expect(clearXlideGlobalAnalysisRuleSeverityOverride(config, 'unknown-call')).resolves.toEqual({
            key: 'xlide.analysis.ruleSeverityOverrides',
            value: { 'option-explicit-missing': 'off' },
            changed: true,
        });
        expect(updates).toEqual([
            {
                key: 'analysis.ruleSeverityOverrides',
                value: {
                    'option-explicit-missing': 'off',
                    'unknown-call': 'warning',
                },
                target: true,
            },
            {
                key: 'analysis.ruleSeverityOverrides',
                value: { 'option-explicit-missing': 'off' },
                target: true,
            },
        ]);
    });

    it('resets global settings through the machine settings target', async () => {
        const updates: Array<{ key: string; value: unknown; target: unknown }> = [];
        const config = fakeConfig({
            pythonPath: 'C:\\Python\\python.exe',
        }, new Set(['pythonPath']), updates);

        await expect(resetXlideGlobalSettingValue(config, 'pythonPath')).resolves.toEqual({
            key: 'xlide.pythonPath',
            value: '',
            changed: true,
        });
        expect(updates).toEqual([{
            key: 'pythonPath',
            value: undefined,
            target: true,
        }]);
    });
});

function fakeConfig(
    values: Record<string, unknown>,
    machineKeys = new Set<string>(),
    updates: Array<{ key: string; value: unknown; target: unknown }> = [],
): vscode.WorkspaceConfiguration {
    return {
        get: (key: string, fallback?: unknown) => key in values ? values[key] : fallback,
        inspect: (key: string) => machineKeys.has(key) ? { globalValue: values[key] } : {},
        update: (key: string, value: unknown, target?: unknown) => {
            if (value === undefined) {
                delete values[key];
                machineKeys.delete(key);
            } else {
                values[key] = value;
                machineKeys.add(key);
            }
            updates.push({ key, value, target });
            return Promise.resolve();
        },
    } as unknown as vscode.WorkspaceConfiguration;
}
