import * as vscode from 'vscode';
import {
    allDiagnosticRuleMetadata,
    allowedDiagnosticSeverityOverridesForCode,
    type DiagnosticRuleMetadata,
    type DiagnosticSeverityOverride,
} from './analyzer/diagnostics/ruleMetadata';
import { ANALYSIS_SEVERITIES, type AnalysisSeverityFilter } from './analysisSettingsCore';
import {
    XLIDE_GLOBAL_SETTING_KEYS,
    clearXlideGlobalAnalysisRuleSeverityOverride,
    resetXlideGlobalSettingValue,
    resolvedXlideGlobalSettingsFromConfig,
    setXlideGlobalAnalysisRuleSeverityOverride,
    setXlideGlobalSettingValue,
    validateXlideGlobalSettingsFromConfig,
    xlideGlobalSettingCards,
    type ResolvedXlideGlobalSetting,
    type XlideGlobalSettingCard,
    type XlideGlobalSettingKey,
    type XlideGlobalSettingSection,
    type XlideGlobalSettingsProblem,
} from './globalSettings';
import { escapeAttr, escapeHtml, randomNonce } from './webview/html';
import { webviewHeadHtml, WEBVIEW_TOAST_CSS } from './webview/page';
import { WEBVIEW_BODY_CSS } from './webview/styles';
import { renderWebviewTemplate } from './webview/templates';
import { registerXlideCommand } from './xlideCommandRegistration';
import { errorMessage } from './util/errors';

interface XlideGlobalSettingsRuleOption {
    code: string;
    title: string;
    category: string;
    defaultSeverity: string;
    diagnosticKind: string;
    source: string;
    allowedSeverityOverrides: readonly DiagnosticSeverityOverride[];
}

interface XlideGlobalSettingsModel {
    settings: ResolvedXlideGlobalSetting<unknown>[];
    problems: XlideGlobalSettingsProblem[];
    rules: XlideGlobalSettingsRuleOption[];
}

interface XlideGlobalSettingsMessage {
    type?: unknown;
    key?: unknown;
    value?: unknown;
    code?: unknown;
    severity?: unknown;
}

const SETTING_KEY_SET = new Set<string>(XLIDE_GLOBAL_SETTING_KEYS);

const SETTING_SECTIONS: ReadonlyArray<{ id: XlideGlobalSettingSection; title: string }> = [
    { id: 'runtime', title: 'Runtime' },
    { id: 'editor', title: 'Editor And Diagnostics' },
    { id: 'docs', title: 'Documentation' },
    { id: 'analysis', title: 'Analysis' },
];

function registerXlideGlobalSettingsWebview(out: vscode.OutputChannel): vscode.Disposable {
    let panel: vscode.WebviewPanel | undefined;

    const render = () => {
        if (!panel) {
            return;
        }
        panel.webview.html = renderXlideGlobalSettingsHtml(
            panel.webview,
            buildXlideGlobalSettingsModel(vscode.workspace.getConfiguration('xlide')),
        );
    };

    const command = registerXlideCommand('xlide.openGlobalSettings', () => {
        if (panel) {
            panel.reveal(vscode.ViewColumn.Active);
            render();
            return;
        }

        panel = vscode.window.createWebviewPanel(
            'xlide.globalSettings',
            'XLIDE Global Settings',
            vscode.ViewColumn.Active,
            {
                // No retainContextWhenHidden: the document is rebuilt on every
                // config change and scroll/search restore via vscode.getState.
                enableScripts: true,
            },
        );
        const panelDisposables: vscode.Disposable[] = [];
        panel.webview.onDidReceiveMessage(async (message: XlideGlobalSettingsMessage) => {
            try {
                // The xlide configuration watcher below is the single render
                // trigger for applied updates; rendering here too would rebuild
                // the document twice per change.
                await applyXlideGlobalSettingsMessage(
                    vscode.workspace.getConfiguration('xlide'),
                    message,
                );
            } catch (err) {
                const error = errorMessage(err);
                out.appendLine(`[globalSettings] Settings update failed: ${error}`);
                await panel?.webview.postMessage({
                    type: 'error',
                    error: `XLIDE: Settings update failed: ${error}`,
                });
                render();
            }
        }, undefined, panelDisposables);
        panel.onDidDispose(() => {
            for (const disposable of panelDisposables) {
                disposable.dispose();
            }
            panel = undefined;
        });
        render();
    });

    const configWatcher = vscode.workspace.onDidChangeConfiguration((event) => {
        if (event.affectsConfiguration('xlide')) {
            render();
        }
    });

    return vscode.Disposable.from(command, configWatcher, {
        dispose: () => panel?.dispose(),
    });
}

function buildXlideGlobalSettingsModel(
    config: vscode.WorkspaceConfiguration,
): XlideGlobalSettingsModel {
    return {
        settings: resolvedXlideGlobalSettingsFromConfig(config),
        problems: validateXlideGlobalSettingsFromConfig(config),
        rules: allDiagnosticRuleMetadata().map(ruleOption),
    };
}

async function applyXlideGlobalSettingsMessage(
    config: vscode.WorkspaceConfiguration,
    message: XlideGlobalSettingsMessage,
): Promise<boolean> {
    if (!message || typeof message !== 'object' || typeof message.type !== 'string') {
        return false;
    }
    if (message.type === 'updateSetting') {
        if (!isXlideGlobalSettingKey(message.key)) {
            return false;
        }
        await setXlideGlobalSettingValue(config, message.key, message.value);
        return true;
    }
    if (message.type === 'resetSetting') {
        if (!isXlideGlobalSettingKey(message.key)) {
            return false;
        }
        await resetXlideGlobalSettingValue(config, message.key);
        return true;
    }
    if (message.type === 'setRuleSeverityOverride') {
        if (typeof message.code !== 'string') {
            return false;
        }
        if (typeof message.severity === 'string' && message.severity.trim().length > 0) {
            await setXlideGlobalAnalysisRuleSeverityOverride(config, message.code, message.severity);
        } else {
            await clearXlideGlobalAnalysisRuleSeverityOverride(config, message.code);
        }
        return true;
    }
    return false;
}

function renderXlideGlobalSettingsHtml(
    webviewOrModel: vscode.Webview | XlideGlobalSettingsModel,
    maybeModel?: XlideGlobalSettingsModel,
): string {
    const model = maybeModel ?? webviewOrModel as XlideGlobalSettingsModel;
    const nonce = randomNonce();
    const settingProblems = settingProblemsByKey(model.problems);
    return renderWebviewTemplate('assets/webview/globalSettings.html', {
        head: webviewHeadHtml(nonce, 'XLIDE Global Settings'),
        nonce,
        css: renderWebviewTemplate('assets/webview/globalSettings.css', {
            bodyCss: WEBVIEW_BODY_CSS,
            toastCss: WEBVIEW_TOAST_CSS,
        }),
        problemBanner: model.problems.length > 0 ? `<div class="problemBanner">${escapeHtml(model.problems.length === 1
        ? model.problems[0].message
        : `${model.problems.length} XLIDE settings need attention.`)}</div>` : '',
        sections: SETTING_SECTIONS.map((section) => renderSettingsSection(model, settingProblems, section)).join('\n        '),
        js: renderWebviewTemplate('assets/webview/globalSettings.js', {}),
    });
}

function renderSettingsSection(
    model: XlideGlobalSettingsModel,
    problems: ReadonlyMap<string, XlideGlobalSettingsProblem[]>,
    section: { id: XlideGlobalSettingSection; title: string },
): string {
    const cards = xlideGlobalSettingCards().filter((card) => card.section === section.id);
    return `<section class="section" aria-label="${escapeAttr(section.title)}">
        <div class="sectionHeader"><h2>${escapeHtml(section.title)}</h2></div>
        <div class="grid">
            ${cards.map((card) => renderSettingControl(model, problems, card)).join('\n            ')}
        </div>
    </section>`;
}

function renderSettingControl(
    model: XlideGlobalSettingsModel,
    problems: ReadonlyMap<string, XlideGlobalSettingsProblem[]>,
    card: XlideGlobalSettingCard,
): string {
    switch (card.control.kind) {
        case 'text':
            return renderTextSetting(model, problems, card.key, card.label);
        case 'boolean':
            return renderBooleanSetting(model, problems, card.key, card.label);
        case 'enum':
            return renderEnumSetting(model, problems, card.key, card.label, card.control.values);
        case 'severityFilter':
            return renderVisibleSeveritiesSetting(model, problems, card.key, card.label);
        case 'rulePicker':
            return renderRulePickerSetting(model, problems, card.key, card.label);
        case 'ruleSeverityOverrides':
            return renderRuleSeverityOverridesSetting(model, problems, card.key, card.label);
    }
}

function renderTextSetting(
    model: XlideGlobalSettingsModel,
    problems: ReadonlyMap<string, XlideGlobalSettingsProblem[]>,
    key: XlideGlobalSettingKey,
    title: string,
): string {
    const value = String(settingValue(model, key, ''));
    return renderSettingCard(model, problems, key, title, `
        <input type="text" data-setting-kind="text" data-setting-key="${escapeAttr(key)}" value="${escapeAttr(value)}" aria-label="${escapeAttr(title)}">
    `);
}

function renderBooleanSetting(
    model: XlideGlobalSettingsModel,
    problems: ReadonlyMap<string, XlideGlobalSettingsProblem[]>,
    key: XlideGlobalSettingKey,
    title: string,
): string {
    const checked = settingValue<boolean>(model, key, false) === true ? ' checked' : '';
    return renderSettingCard(model, problems, key, title, `
        <label class="checkRow">
            <input type="checkbox" data-setting-kind="boolean" data-setting-key="${escapeAttr(key)}"${checked}>
            <span>${escapeHtml(title)}</span>
        </label>
    `);
}

function renderEnumSetting(
    model: XlideGlobalSettingsModel,
    problems: ReadonlyMap<string, XlideGlobalSettingsProblem[]>,
    key: XlideGlobalSettingKey,
    title: string,
    values: readonly string[],
): string {
    const current = String(settingValue(model, key, values[0] ?? ''));
    return renderSettingCard(model, problems, key, title, `
        <select data-setting-kind="enum" data-setting-key="${escapeAttr(key)}" aria-label="${escapeAttr(title)}">
            ${values.map((value) => `<option value="${escapeAttr(value)}"${value === current ? ' selected' : ''}>${escapeHtml(titleCase(value))}</option>`).join('')}
        </select>
    `);
}

function renderVisibleSeveritiesSetting(
    model: XlideGlobalSettingsModel,
    problems: ReadonlyMap<string, XlideGlobalSettingsProblem[]>,
    key: XlideGlobalSettingKey,
    title: string,
): string {
    const selected = new Set(settingValue<AnalysisSeverityFilter[]>(model, key, []));
    return renderSettingCard(model, problems, key, title, `
        <div class="severitySet">
            ${ANALYSIS_SEVERITIES.map((severity) => `<label class="severityRow">
                <input type="checkbox" data-severity-filter value="${escapeAttr(severity)}"${selected.has(severity) ? ' checked' : ''}>
                <span>${escapeHtml(titleCase(severity))}</span>
            </label>`).join('')}
        </div>
    `);
}

function renderRulePickerSetting(
    model: XlideGlobalSettingsModel,
    problems: ReadonlyMap<string, XlideGlobalSettingsProblem[]>,
    key: XlideGlobalSettingKey,
    title: string,
): string {
    const untrackedRules = new Set(settingValue<string[]>(model, key, []));
    const rows = model.rules.map((rule) => {
        const checked = untrackedRules.has(rule.code) ? ' checked' : '';
        const search = `${rule.title} ${rule.code} ${rule.category} ${rule.defaultSeverity}`.toLowerCase();
        return `<label class="ruleRow" data-rule-row data-search="${escapeAttr(search)}">
            <input type="checkbox" data-rule-untracked value="${escapeAttr(rule.code)}"${checked}>
            <span>
                <span class="ruleTitle">${escapeHtml(rule.title)}</span>
                <span class="code">${escapeHtml(rule.code)}</span>
            </span>
            <span class="tagSet">
                <span class="tag">${escapeHtml(rule.defaultSeverity)}</span>
                <span class="tag">${escapeHtml(rule.category)}</span>
            </span>
        </label>`;
    }).join('');
    return renderSettingCard(model, problems, key, title, `
        <div class="ruleTools">
            <input type="search" id="ruleSearch" aria-label="Search Analysis Rules" placeholder="Search Rules">
        </div>
        <div class="ruleList" id="ruleList" aria-label="${escapeAttr(title)}">
            ${rows || '<div class="empty">No rules</div>'}
        </div>
    `, 'wide');
}

function renderRuleSeverityOverridesSetting(
    model: XlideGlobalSettingsModel,
    problems: ReadonlyMap<string, XlideGlobalSettingsProblem[]>,
    key: XlideGlobalSettingKey,
    title: string,
): string {
    const overrides = settingValue<Record<string, DiagnosticSeverityOverride>>(model, key, {});
    const rules = model.rules
        .filter((rule) => rule.allowedSeverityOverrides.length > 0)
        .sort((left, right) => left.title.localeCompare(right.title) || left.code.localeCompare(right.code));
    const rows = rules.map((rule) => `<div class="overrideRow">
        <div>
            <div class="ruleTitle">${escapeHtml(rule.title)}</div>
            <div class="code">${escapeHtml(rule.code)}</div>
        </div>
        <select data-rule-severity data-rule-code="${escapeAttr(rule.code)}" aria-label="${escapeAttr(`${rule.title} Severity Override`)}">
            <option value=""${overrides[rule.code] ? '' : ' selected'}>Default (${escapeHtml(rule.defaultSeverity)})</option>
            ${rule.allowedSeverityOverrides.map((severity) => `<option value="${escapeAttr(severity)}"${overrides[rule.code] === severity ? ' selected' : ''}>${escapeHtml(titleCase(severity))}</option>`).join('')}
        </select>
    </div>`).join('');
    return renderSettingCard(model, problems, key, title, `
        <div class="overrideList" id="overrideList">
            ${rows || '<div class="empty">No severity overrides available</div>'}
        </div>
    `, 'wide');
}

function renderSettingCard(
    model: XlideGlobalSettingsModel,
    problems: ReadonlyMap<string, XlideGlobalSettingsProblem[]>,
    key: XlideGlobalSettingKey,
    title: string,
    body: string,
    extraClass = '',
): string {
    const fullKey = `xlide.${key}`;
    const cardProblems = problems.get(fullKey) ?? [];
    return `<section class="card ${escapeAttr(extraClass)}" data-setting-card="${escapeAttr(fullKey)}">
        <div class="cardHeader">
            <div class="titleBlock">
                <h3>${escapeHtml(title)}</h3>
                <div class="source">Source: ${escapeHtml(sourceLabel(settingSource(model, key)))}</div>
            </div>
            <button type="button" data-reset-setting="${escapeAttr(key)}">Reset</button>
        </div>
        <div class="cardBody">
            ${body}
            ${cardProblems.map((problem) => `<div class="problem">${escapeHtml(problem.message)}</div>`).join('')}
        </div>
    </section>`;
}

function ruleOption(rule: DiagnosticRuleMetadata): XlideGlobalSettingsRuleOption {
    return {
        code: rule.code,
        title: rule.title,
        category: rule.category,
        defaultSeverity: rule.defaultSeverity,
        diagnosticKind: rule.diagnosticKind,
        source: rule.source,
        allowedSeverityOverrides: allowedDiagnosticSeverityOverridesForCode(rule.code),
    };
}

function settingValue<T>(
    model: XlideGlobalSettingsModel,
    key: XlideGlobalSettingKey,
    fallback: T,
): T {
    const setting = model.settings.find((entry) => entry.key === `xlide.${key}`);
    return setting ? setting.value as T : fallback;
}

function settingSource(
    model: XlideGlobalSettingsModel,
    key: XlideGlobalSettingKey,
): ResolvedXlideGlobalSetting<unknown>['source'] {
    return model.settings.find((entry) => entry.key === `xlide.${key}`)?.source ?? 'unknown';
}

function settingProblemsByKey(
    problems: readonly XlideGlobalSettingsProblem[],
): ReadonlyMap<string, XlideGlobalSettingsProblem[]> {
    const byKey = new Map<string, XlideGlobalSettingsProblem[]>();
    for (const problem of problems) {
        const existing = byKey.get(problem.key) ?? [];
        existing.push(problem);
        byKey.set(problem.key, existing);
    }
    return byKey;
}

function sourceLabel(source: ResolvedXlideGlobalSetting<unknown>['source']): string {
    switch (source) {
        case 'machine':
            return 'Machine';
        case 'default':
            return 'Default';
        default:
            return 'Unknown';
    }
}

function isXlideGlobalSettingKey(value: unknown): value is XlideGlobalSettingKey {
    return typeof value === 'string' && SETTING_KEY_SET.has(value);
}

function titleCase(value: string): string {
    return value
        .split(/[\s.-]+/)
        .filter((part) => part.length > 0)
        .map((part) => part.charAt(0).toUpperCase() + part.slice(1))
        .join(' ');
}

export {
    applyXlideGlobalSettingsMessage,
    buildXlideGlobalSettingsModel,
    registerXlideGlobalSettingsWebview,
    renderXlideGlobalSettingsHtml,
    type XlideGlobalSettingsModel,
    type XlideGlobalSettingsRuleOption,
};
