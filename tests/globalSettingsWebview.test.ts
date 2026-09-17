import { describe, expect, it, vi } from 'vitest';
import * as vscode from 'vscode';
import { fakeConfig } from './helpers/fakeConfig';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());

import {
    applyXlideGlobalSettingsMessage,
    buildXlideGlobalSettingsModel,
    registerXlideGlobalSettingsWebview,
    renderXlideGlobalSettingsHtml,
} from '../src/globalSettingsWebview';

const validSettings = {
    'officeIntegration.attachToRunning': true,
    'officeIntegration.coordinationMode': 'block',
    'officeIntegration.trackOpenedFiles': true,
    'officeIntegration.reopenAfterClose': true,
    'officeIntegration.reopenMode': 'readOnly',
    'officeIntegration.reopenReadOnlyAfterSave': false,
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
        expect(html).toContain('id="toast"');
        expect(html).toContain('showToast(event.data.error');
        expect(html).not.toContain('Add Item');
    });

    it('renders info bubbles carrying each setting description as hover help', () => {
        const html = renderXlideGlobalSettingsHtml(buildXlideGlobalSettingsModel(fakeConfig(validSettings)));

        expect(html).toContain('class="infoBubble"');
        // the renamed coordination-mode label and its description tooltip
        expect(html).toContain('When a File is Open in Its Application');
        expect(html).toContain('Auto Expand And Collapse Explorer Tree');
        expect(html).toContain('What XLIDE does when Excel, Word, PowerPoint or Access holds the file open');
        // The section serves every Office application, and says so.
        expect(html).toContain('Office Integration');
        expect(html).not.toContain('Excel Integration');
        // the reopen-as control offers the spaced "Last State" option
        expect(html).toContain('>Last State<');
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
            'analysis.untrackedRules': [''],
        }, new Set(['analysis.untrackedRules'])));
        const html = renderXlideGlobalSettingsHtml(model);

        expect(model.problems).toEqual([{
            key: 'xlide.analysis.untrackedRules',
            message: 'Expected "xlide.analysis.untrackedRules" entries to be non-empty strings.',
            severity: 'warning',
        }]);
        expect(html).toContain('non-empty strings');
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

    it('surfaces settings update failures to the panel and the output channel', async () => {
        const failure = new Error('settings.json is read-only');
        const config = {
            ...fakeConfig(validSettings),
            update: () => Promise.reject(failure),
        } as unknown as vscode.WorkspaceConfiguration;
        vi.mocked(vscode.workspace.getConfiguration).mockReturnValue(config);
        let messageHandler: ((message: unknown) => Promise<void>) | undefined;
        const postMessage = vi.fn(() => Promise.resolve(true));
        const panel = {
            webview: {
                html: '',
                onDidReceiveMessage: vi.fn((handler: (message: unknown) => Promise<void>) => {
                    messageHandler = handler;
                    return { dispose: vi.fn() };
                }),
                postMessage,
            },
            onDidDispose: vi.fn(),
            reveal: vi.fn(),
            dispose: vi.fn(),
        };
        vi.mocked(vscode.window.createWebviewPanel).mockReturnValue(panel as unknown as vscode.WebviewPanel);
        let openCommand: (() => unknown) | undefined;
        vi.mocked(vscode.commands.registerCommand).mockImplementation((command, callback) => {
            if (command === 'xlide.openGlobalSettings') {
                openCommand = callback as () => unknown;
            }
            return { dispose: vi.fn() } as unknown as vscode.Disposable;
        });
        const appendLine = vi.fn();
        const registration = registerXlideGlobalSettingsWebview(
            { appendLine } as unknown as vscode.OutputChannel,
        );

        await openCommand?.();
        panel.webview.html = '';
        await messageHandler?.({ type: 'updateSetting', key: 'docs.enabled', value: false });

        expect(appendLine).toHaveBeenCalledWith(expect.stringContaining('settings.json is read-only'));
        expect(postMessage).toHaveBeenCalledWith({
            type: 'error',
            error: expect.stringContaining('settings.json is read-only'),
        });
        expect(panel.webview.html).toContain('XLIDE Global Settings');
        registration.dispose();
    });
});
