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
    'explorer.autoExpandCollapse': true,
    'explorer.view': 'tree',
    'docs.enabled': true,
    'docs.metadataGlob': '**/*.vbref.xml',
    'performance.trace': false,
};

describe('globalSettings', () => {
    it('accepts the extension default settings shape', () => {
        expect(validateXlideGlobalSettingsValues(validSettings)).toEqual([]);
    });

    it('never reports settings that are unset (the version-upgrade case)', () => {
        // A newly-added setting the user has not configured must produce no
        // problem - validation only sees values that are actually present, so an
        // upgrade that adds settings never blasts "values not set correctly".
        const withoutOne = { ...validSettings };
        delete (withoutOne as Record<string, unknown>)['analysis.visibleSeverities'];
        expect(validateXlideGlobalSettingsValues(withoutOne)).toEqual([]);
        // A brand-new install with nothing configured reports nothing.
        expect(validateXlideGlobalSettingsValues({})).toEqual([]);
    });

    it('reports structurally malformed analysis settings but tolerates unknown rule codes', () => {
        expect(validateXlideGlobalSettingsValues({
            ...validSettings,
            'analysis.visibleSeverities': ['error', 'hint'],
            // 'not-a-rule' is a well-formed but unknown/renamed code - tolerated, not reported.
            'analysis.untrackedRules': ['option-explicit-missing', 'not-a-rule'],
            'analysis.ruleSeverityOverrides': {
                'option-explicit-missing': 'error', // known code, disallowed severity -> reported
                'not-a-rule': 'warning',            // unknown code -> tolerated, not reported
            },
        })).toEqual([
            {
                key: 'xlide.analysis.ruleSeverityOverrides',
                message: 'Expected "xlide.analysis.ruleSeverityOverrides.option-explicit-missing" to be one of: off.',
                severity: 'warning',
            },
            {
                key: 'xlide.analysis.visibleSeverities',
                message: 'Expected "xlide.analysis.visibleSeverities" entries to be one of: error, warning, information.',
                severity: 'warning',
            },
        ]);
    });

    it('never warns on stale/unknown rule codes (the post-upgrade case)', () => {
        // A code the user set in an older version that has since been renamed or
        // removed must not warn - every apply path drops it silently.
        expect(validateXlideGlobalSettingsValues({
            ...validSettings,
            'analysis.untrackedRules': ['since-removed-rule'],
            'analysis.ruleSeverityOverrides': { 'since-renamed-rule': 'off' },
        })).toEqual([]);
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
            'docs.metadataGlob': 'refs/**/*.vbref.xml',
        }, new Set(['docs.metadataGlob'])));

        expect(settings.map((setting) => setting.key)).toEqual([
            'xlide.agent.showWriteDiffs',
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
            'xlide.explorer.autoExpandCollapse',
            'xlide.explorer.view',
            'xlide.formRun.injectShowMacro',
            'xlide.performance.trace',
        ]);
        expect(settings.find((setting) => setting.key === 'xlide.docs.metadataGlob')).toMatchObject({
            value: 'refs/**/*.vbref.xml',
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
            'docs.metadataGlob': 'refs/**/*.vbref.xml',
        }, new Set(['docs.metadataGlob']), updates);

        await expect(resetXlideGlobalSettingValue(config, 'docs.metadataGlob')).resolves.toEqual({
            key: 'xlide.docs.metadataGlob',
            value: '**/*.vbref.xml',
            changed: true,
        });
        expect(updates).toEqual([{
            key: 'docs.metadataGlob',
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
