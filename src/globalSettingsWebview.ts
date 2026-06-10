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
    type ResolvedXlideGlobalSetting,
    type XlideGlobalSettingKey,
    type XlideGlobalSettingsProblem,
} from './globalSettings';
import { escapeAttr, escapeHtml, randomNonce } from './webview/html';
import { registerXlideCommand } from './xlideCommandRegistration';

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
                enableScripts: true,
                retainContextWhenHidden: true,
            },
        );
        const panelDisposables: vscode.Disposable[] = [];
        panel.webview.onDidReceiveMessage(async (message: XlideGlobalSettingsMessage) => {
            try {
                const applied = await applyXlideGlobalSettingsMessage(
                    vscode.workspace.getConfiguration('xlide'),
                    message,
                );
                if (applied) {
                    render();
                }
            } catch (err) {
                const error = err instanceof Error ? err.message : String(err);
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
    return `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'nonce-${nonce}'; script-src 'nonce-${nonce}';">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>XLIDE Global Settings</title>
    <style nonce="${nonce}">
        :root {
            color-scheme: light dark;
        }
        * {
            box-sizing: border-box;
        }
        body {
            margin: 0;
            color: var(--vscode-foreground);
            background: var(--vscode-editor-background);
            font-family: var(--vscode-font-family);
            font-size: var(--vscode-font-size);
            line-height: 1.4;
        }
        .shell {
            max-width: 1120px;
            margin: 0 auto;
            padding: 24px;
        }
        .pageHeader {
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            gap: 16px;
            padding-bottom: 18px;
            border-bottom: 1px solid var(--vscode-panel-border);
        }
        h1,
        h2,
        h3 {
            margin: 0;
            letter-spacing: 0;
        }
        h1 {
            font-size: 24px;
            line-height: 1.2;
        }
        h2 {
            font-size: 17px;
            font-weight: 700;
        }
        h3 {
            font-size: 15px;
            font-weight: 700;
        }
        .subtitle,
        .source,
        .description,
        .code,
        .empty {
            color: var(--vscode-descriptionForeground);
        }
        .subtitle {
            margin-top: 4px;
        }
        .section {
            margin-top: 22px;
        }
        .sectionHeader {
            display: flex;
            align-items: baseline;
            justify-content: space-between;
            gap: 12px;
            margin-bottom: 10px;
            padding-bottom: 6px;
            border-bottom: 1px solid var(--vscode-panel-border);
        }
        .grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
            gap: 12px;
        }
        .card {
            border: 1px solid var(--vscode-panel-border);
            border-radius: 6px;
            background: var(--vscode-sideBar-background);
            overflow: hidden;
            min-width: 0;
        }
        .wide {
            grid-column: 1 / -1;
        }
        .cardHeader {
            display: flex;
            justify-content: space-between;
            gap: 12px;
            align-items: flex-start;
            padding: 12px 14px;
            border-bottom: 1px solid var(--vscode-panel-border);
            background: var(--vscode-editorWidget-background, var(--vscode-editor-background));
        }
        .cardBody {
            padding: 14px;
        }
        .titleBlock {
            min-width: 0;
        }
        .source {
            margin-top: 2px;
            font-size: 12px;
        }
        .problem {
            margin-top: 10px;
            border-left: 3px solid var(--vscode-editorError-foreground);
            padding-left: 9px;
            color: var(--vscode-editorError-foreground);
        }
        .problemBanner {
            margin-top: 14px;
            border: 1px solid var(--vscode-editorError-foreground);
            border-radius: 6px;
            padding: 10px 12px;
            color: var(--vscode-editorError-foreground);
            background: color-mix(in srgb, var(--vscode-editorError-foreground) 8%, transparent);
        }
        input[type="text"],
        input[type="search"],
        select {
            width: 100%;
            min-width: 0;
            border: 1px solid var(--vscode-input-border, var(--vscode-panel-border));
            border-radius: 4px;
            padding: 6px 8px;
            color: var(--vscode-input-foreground);
            background: var(--vscode-input-background);
            font: inherit;
        }
        input[type="checkbox"] {
            width: 16px;
            height: 16px;
            margin: 0;
        }
        input:focus,
        select:focus,
        button:focus {
            outline: 1px solid var(--vscode-focusBorder);
            outline-offset: 2px;
        }
        button {
            border: 1px solid var(--vscode-button-border, transparent);
            border-radius: 4px;
            padding: 4px 8px;
            color: var(--vscode-button-secondaryForeground);
            background: var(--vscode-button-secondaryBackground);
            font: inherit;
            cursor: pointer;
            min-height: 26px;
        }
        button:hover {
            background: var(--vscode-button-secondaryHoverBackground);
        }
        .checkRow,
        .severityRow,
        .ruleRow,
        .overrideRow {
            display: grid;
            gap: 10px;
            align-items: center;
        }
        .checkRow {
            grid-template-columns: 18px minmax(0, 1fr);
        }
        .severitySet {
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
        }
        .severityRow {
            grid-template-columns: 18px auto;
        }
        .ruleTools {
            margin-bottom: 10px;
        }
        .ruleList,
        .overrideList {
            display: grid;
            gap: 8px;
            max-height: 430px;
            overflow: auto;
            padding-right: 4px;
        }
        .ruleRow {
            grid-template-columns: 18px minmax(0, 1fr) auto;
            border: 1px solid var(--vscode-panel-border);
            border-radius: 4px;
            padding: 8px;
            background: var(--vscode-editor-background);
        }
        .ruleRow[hidden] {
            display: none;
        }
        .ruleTitle {
            font-weight: 600;
            overflow-wrap: anywhere;
        }
        .code {
            font-family: var(--vscode-editor-font-family);
            font-size: 12px;
            overflow-wrap: anywhere;
        }
        .tagSet {
            display: flex;
            gap: 6px;
            flex-wrap: wrap;
            justify-content: flex-end;
        }
        .tag {
            border: 1px solid var(--vscode-panel-border);
            border-radius: 4px;
            padding: 2px 6px;
            color: var(--vscode-foreground);
            background: var(--vscode-editorWidget-background, var(--vscode-input-background));
            font-size: 12px;
            white-space: nowrap;
        }
        .overrideRow {
            grid-template-columns: minmax(0, 1fr) minmax(140px, 220px);
            border-bottom: 1px solid var(--vscode-panel-border);
            padding: 8px 0;
        }
        .overrideRow:last-child {
            border-bottom: 0;
        }
        .toast {
            position: fixed;
            right: 18px;
            bottom: 18px;
            max-width: 360px;
            border: 1px solid var(--vscode-panel-border);
            border-radius: 6px;
            padding: 10px 12px;
            color: var(--vscode-notifications-foreground);
            background: var(--vscode-notifications-background);
            box-shadow: 0 8px 22px rgba(0, 0, 0, 0.32);
            opacity: 0;
            transform: translateY(8px);
            transition: opacity 120ms ease, transform 120ms ease;
            pointer-events: none;
        }
        .toast.visible {
            opacity: 1;
            transform: translateY(0);
        }
        @media (max-width: 640px) {
            .shell {
                padding: 16px;
            }
            .pageHeader,
            .sectionHeader,
            .cardHeader {
                display: grid;
            }
            .overrideRow,
            .ruleRow {
                grid-template-columns: 18px minmax(0, 1fr);
            }
            .overrideRow {
                grid-template-columns: minmax(0, 1fr);
            }
            .tagSet {
                grid-column: 2;
                justify-content: flex-start;
            }
        }
    </style>
</head>
<body>
    <main class="shell">
        <header class="pageHeader">
            <div>
                <h1>XLIDE Global Settings</h1>
                <div class="subtitle">VS Code / Machine</div>
            </div>
        </header>
        ${model.problems.length > 0 ? `<div class="problemBanner">${escapeHtml(model.problems.length === 1
        ? model.problems[0].message
        : `${model.problems.length} XLIDE settings need attention.`)}</div>` : ''}
        ${renderRuntimeSection(model, settingProblems)}
        ${renderEditorSection(model, settingProblems)}
        ${renderDocsSection(model, settingProblems)}
        ${renderAnalysisSection(model, settingProblems)}
    </main>
    <div class="toast" id="toast"></div>
    <script nonce="${nonce}">
        const vscode = acquireVsCodeApi();
        const toast = document.getElementById('toast');
        let toastTimer;

        restoreGlobalSettingsState();

        function showToast(message) {
            toast.textContent = message;
            toast.classList.add('visible');
            clearTimeout(toastTimer);
            toastTimer = setTimeout(() => toast.classList.remove('visible'), 2600);
        }

        window.addEventListener('message', (event) => {
            if (event.data?.type === 'error') {
                showToast(event.data.error || 'XLIDE settings update failed');
            }
        });

        document.addEventListener('change', (event) => {
            const target = event.target;
            if (!(target instanceof HTMLElement)) {
                return;
            }
            if (target.matches('input[data-setting-kind="boolean"]')) {
                persistGlobalSettingsState();
                vscode.postMessage({
                    type: 'updateSetting',
                    key: target.dataset.settingKey,
                    value: target.checked === true
                });
                return;
            }
            if (target.matches('input[data-setting-kind="text"]')) {
                persistGlobalSettingsState();
                vscode.postMessage({
                    type: 'updateSetting',
                    key: target.dataset.settingKey,
                    value: target.value
                });
                return;
            }
            if (target.matches('select[data-setting-kind="enum"]')) {
                persistGlobalSettingsState();
                vscode.postMessage({
                    type: 'updateSetting',
                    key: target.dataset.settingKey,
                    value: target.value
                });
                return;
            }
            if (target.matches('input[data-severity-filter]')) {
                persistGlobalSettingsState();
                vscode.postMessage({
                    type: 'updateSetting',
                    key: 'analysis.visibleSeverities',
                    value: checkedValues('input[data-severity-filter]:checked')
                });
                return;
            }
            if (target.matches('input[data-rule-untracked]')) {
                persistGlobalSettingsState();
                vscode.postMessage({
                    type: 'updateSetting',
                    key: 'analysis.untrackedRules',
                    value: checkedValues('input[data-rule-untracked]:checked')
                });
                return;
            }
            if (target.matches('select[data-rule-severity]')) {
                persistGlobalSettingsState();
                vscode.postMessage({
                    type: 'setRuleSeverityOverride',
                    code: target.dataset.ruleCode,
                    severity: target.value
                });
            }
        });

        document.addEventListener('input', (event) => {
            const search = event.target.closest?.('#ruleSearch');
            if (!search) {
                return;
            }
            applyRuleSearch(search.value);
            persistGlobalSettingsState();
        });

        document.addEventListener('click', (event) => {
            const button = event.target.closest?.('button[data-reset-setting]');
            if (!button) {
                return;
            }
            persistGlobalSettingsState();
            vscode.postMessage({
                type: 'resetSetting',
                key: button.dataset.resetSetting
            });
        });

        function checkedValues(selector) {
            return Array.from(document.querySelectorAll(selector)).map((input) => input.value);
        }

        function applyRuleSearch(value) {
            const query = value.trim().toLowerCase();
            for (const row of document.querySelectorAll('[data-rule-row]')) {
                row.hidden = query.length > 0 && !row.dataset.search.includes(query);
            }
        }

        function persistGlobalSettingsState() {
            vscode.setState({
                pageScrollTop: window.scrollY,
                ruleSearch: document.getElementById('ruleSearch')?.value ?? '',
                ruleListScrollTop: document.getElementById('ruleList')?.scrollTop ?? 0,
                overrideListScrollTop: document.getElementById('overrideList')?.scrollTop ?? 0
            });
        }

        function restoreGlobalSettingsState() {
            const state = typeof vscode.getState === 'function' ? vscode.getState() || {} : {};
            const ruleSearch = document.getElementById('ruleSearch');
            if (ruleSearch && typeof state.ruleSearch === 'string') {
                ruleSearch.value = state.ruleSearch;
                applyRuleSearch(state.ruleSearch);
            }
            requestAnimationFrame(() => {
                if (typeof state.pageScrollTop === 'number') {
                    window.scrollTo(0, state.pageScrollTop);
                }
                const ruleList = document.getElementById('ruleList');
                if (ruleList && typeof state.ruleListScrollTop === 'number') {
                    ruleList.scrollTop = state.ruleListScrollTop;
                }
                const overrideList = document.getElementById('overrideList');
                if (overrideList && typeof state.overrideListScrollTop === 'number') {
                    overrideList.scrollTop = state.overrideListScrollTop;
                }
            });
        }
    </script>
</body>
</html>`;
}

function renderRuntimeSection(
    model: XlideGlobalSettingsModel,
    problems: ReadonlyMap<string, XlideGlobalSettingsProblem[]>,
): string {
    return `<section class="section" aria-label="Runtime">
        <div class="sectionHeader"><h2>Runtime</h2></div>
        <div class="grid">
            ${renderTextSetting(model, problems, 'pythonPath', 'Python Path')}
            ${renderBooleanSetting(model, problems, 'attachToRunningExcel', 'Attach To Running Excel')}
        </div>
    </section>`;
}

function renderEditorSection(
    model: XlideGlobalSettingsModel,
    problems: ReadonlyMap<string, XlideGlobalSettingsProblem[]>,
): string {
    return `<section class="section" aria-label="Editor And Diagnostics">
        <div class="sectionHeader"><h2>Editor And Diagnostics</h2></div>
        <div class="grid">
            ${renderBooleanSetting(model, problems, 'diagnostics.enabled', 'Diagnostics Enabled')}
            ${renderEnumSetting(model, problems, 'editor.blockLayout', 'Editor Block Layout', ['comfy', 'compact'])}
        </div>
    </section>`;
}

function renderDocsSection(
    model: XlideGlobalSettingsModel,
    problems: ReadonlyMap<string, XlideGlobalSettingsProblem[]>,
): string {
    return `<section class="section" aria-label="Documentation">
        <div class="sectionHeader"><h2>Documentation</h2></div>
        <div class="grid">
            ${renderBooleanSetting(model, problems, 'docs.enabled', 'Docs Enabled')}
            ${renderTextSetting(model, problems, 'docs.metadataGlob', 'Docs Metadata Glob')}
        </div>
    </section>`;
}

function renderAnalysisSection(
    model: XlideGlobalSettingsModel,
    problems: ReadonlyMap<string, XlideGlobalSettingsProblem[]>,
): string {
    return `<section class="section" aria-label="Analysis">
        <div class="sectionHeader"><h2>Analysis</h2></div>
        <div class="grid">
            ${renderVisibleSeveritiesSetting(model, problems)}
            ${renderRulePickerSetting(model, problems)}
            ${renderRuleSeverityOverridesSetting(model, problems)}
        </div>
    </section>`;
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
): string {
    const selected = new Set(settingValue<AnalysisSeverityFilter[]>(model, 'analysis.visibleSeverities', []));
    return renderSettingCard(model, problems, 'analysis.visibleSeverities', 'Visible Severities', `
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
): string {
    const untrackedRules = new Set(settingValue<string[]>(model, 'analysis.untrackedRules', []));
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
    return renderSettingCard(model, problems, 'analysis.untrackedRules', 'Globally Untracked Rules', `
        <div class="ruleTools">
            <input type="search" id="ruleSearch" aria-label="Search Analysis Rules" placeholder="Search Rules">
        </div>
        <div class="ruleList" id="ruleList" aria-label="Globally Untracked Rules">
            ${rows || '<div class="empty">No rules</div>'}
        </div>
    `, 'wide');
}

function renderRuleSeverityOverridesSetting(
    model: XlideGlobalSettingsModel,
    problems: ReadonlyMap<string, XlideGlobalSettingsProblem[]>,
): string {
    const overrides = settingValue<Record<string, DiagnosticSeverityOverride>>(
        model,
        'analysis.ruleSeverityOverrides',
        {},
    );
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
    return renderSettingCard(model, problems, 'analysis.ruleSeverityOverrides', 'Rule Severity Overrides', `
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
