import { describe, expect, it, vi } from 'vitest';
import type * as vscode from 'vscode';

vi.mock('vscode', () => ({
    Disposable: {
        from: (...disposables: Array<{ dispose?: () => void }>) => ({
            dispose: () => disposables.forEach((disposable) => disposable.dispose?.()),
        }),
    },
    ViewColumn: { Active: 1 },
    commands: {
        registerCommand: vi.fn(() => ({ dispose: vi.fn() })),
    },
    window: {
        createWebviewPanel: vi.fn(),
    },
    workspace: {
        getConfiguration: vi.fn(),
        onDidChangeConfiguration: vi.fn(() => ({ dispose: vi.fn() })),
    },
}));

import {
    applyXlideGlobalSettingsMessage,
    buildXlideGlobalSettingsModel,
    renderXlideGlobalSettingsHtml,
} from '../src/globalSettingsWebview';

const validSettings = {
    pythonPath: '',
    attachToRunningExcel: true,
    'diagnostics.enabled': true,
    'analysis.ruleSeverityOverrides': {},
    'analysis.visibleSeverities': ['error', 'warning', 'information'],
    'analysis.untrackedRules': [],
    'editor.blockLayout': 'comfy',
    'docs.enabled': true,
    'docs.metadataGlob': '**/*.vbref.xml',
    'performance.trace': false,
};

describe('globalSettingsWebview', () => {
    it('renders global settings with provenance and a searchable known-rule picker', () => {
        const model = buildXlideGlobalSettingsModel(fakeConfig({
            ...validSettings,
            'analysis.untrackedRules': ['option-explicit-missing'],
            'docs.metadataGlob': 'custom/*.vbref.xml',
        }, new Set(['analysis.untrackedRules'])));
        const html = renderXlideGlobalSettingsHtml(model);

        expect(model.settings.find((setting) => setting.key === 'xlide.analysis.untrackedRules'))
            .toMatchObject({
                value: ['option-explicit-missing'],
                source: 'machine',
            });
        expect(model.rules.map((rule) => rule.code)).toContain('option-explicit-missing');
        expect(html).toContain('XLIDE Global Settings');
        expect(html).toContain('VS Code / Machine');
        expect(html).toContain('Globally Untracked Rules');
        expect(html).toContain('id="ruleSearch"');
        expect(html).toContain('id="ruleList"');
        expect(html).toContain('id="overrideList"');
        expect(html).toContain('data-rule-untracked');
        expect(html).toContain('Option Explicit is not specified');
        expect(html).toContain('option-explicit-missing');
        expect(html).toContain('Source: Machine');
        expect(html).toContain('color: var(--vscode-foreground)');
        expect(html).toContain('background: var(--vscode-editorWidget-background, var(--vscode-input-background))');
        expect(html).toContain('persistGlobalSettingsState();');
        expect(html).toContain('restoreGlobalSettingsState();');
        expect(html).toContain('overrideListScrollTop');
        expect(html).not.toContain('Add Item');
    });

    it('sorts rule severity overrides by human rule title', () => {
        const html = renderXlideGlobalSettingsHtml(buildXlideGlobalSettingsModel(fakeConfig(validSettings)));
        const overridesHtml = html.slice(html.indexOf('Rule Severity Overrides'));

        expect(overridesHtml.indexOf('Argument type mismatch')).toBeLessThan(
            overridesHtml.indexOf('Invalid XLIDE analysis suppression directive'),
        );
        expect(overridesHtml.indexOf('Assignment type mismatch')).toBeLessThan(
            overridesHtml.indexOf('Invalid XLIDE analysis suppression directive'),
        );
    });

    it('surfaces validation problems for malformed global analysis settings', () => {
        const model = buildXlideGlobalSettingsModel(fakeConfig({
            ...validSettings,
            'analysis.untrackedRules': ['not-a-rule'],
        }, new Set(['analysis.untrackedRules'])));
        const html = renderXlideGlobalSettingsHtml(model);

        expect(model.problems).toEqual([{
            key: 'xlide.analysis.untrackedRules',
            message: 'Expected "xlide.analysis.untrackedRules" entries to be known analysis rule codes.',
            severity: 'error',
        }]);
        expect(html).toContain('known analysis rule codes');
        expect(html).toContain('data-setting-card="xlide.analysis.untrackedRules"');
    });

    it('applies webview messages through normalized machine-scoped setting updates', async () => {
        const updates: Array<{ key: string; value: unknown; target: unknown }> = [];
        const config = fakeConfig({
            'docs.enabled': false,
        }, new Set(['docs.enabled']), updates);

        await expect(applyXlideGlobalSettingsMessage(config, {
            type: 'updateSetting',
            key: 'analysis.untrackedRules',
            value: ['not-a-rule', ' Option-Explicit-Missing ', 'argument-count'],
        })).resolves.toBe(true);
        await expect(applyXlideGlobalSettingsMessage(config, {
            type: 'setRuleSeverityOverride',
            code: 'option-explicit-missing',
            severity: 'off',
        })).resolves.toBe(true);
        await expect(applyXlideGlobalSettingsMessage(config, {
            type: 'setRuleSeverityOverride',
            code: 'option-explicit-missing',
            severity: '',
        })).resolves.toBe(true);
        await expect(applyXlideGlobalSettingsMessage(config, {
            type: 'resetSetting',
            key: 'docs.enabled',
        })).resolves.toBe(true);
        await expect(applyXlideGlobalSettingsMessage(config, {
            type: 'updateSetting',
            key: 'not.aSetting',
            value: true,
        })).resolves.toBe(false);

        expect(updates).toEqual([
            {
                key: 'analysis.untrackedRules',
                value: ['argument-count', 'option-explicit-missing'],
                target: true,
            },
            {
                key: 'analysis.ruleSeverityOverrides',
                value: { 'option-explicit-missing': 'off' },
                target: true,
            },
            {
                key: 'analysis.ruleSeverityOverrides',
                value: {},
                target: true,
            },
            {
                key: 'docs.enabled',
                value: undefined,
                target: true,
            },
        ]);
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
