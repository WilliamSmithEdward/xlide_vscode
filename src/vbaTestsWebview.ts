import * as path from 'path';
import * as vscode from 'vscode';

export type VbaTestSupportState = 'installed' | 'missing' | 'outdated' | 'blocked' | 'unknown';

export interface VbaTestSupportStatusModel {
    state: VbaTestSupportState;
    title: string;
    description: string;
    actionLabel: string;
    canInstall: boolean;
    canRun: boolean;
}

export type VbaTestRuntimeState = 'installed' | 'missing' | 'blocked' | 'unknown';

export interface VbaTestRuntimeStatusModel {
    state: VbaTestRuntimeState;
    title: string;
    description: string;
    canRun: boolean;
}

export interface VbaTestsPanelModel {
    filePath: string;
    workbookName: string;
    support: VbaTestSupportStatusModel;
    runtime: VbaTestRuntimeStatusModel;
}

export interface VbaTestsPanelOptions {
    getModel: () => Promise<VbaTestsPanelModel>;
    onInstallSupport?: () => Promise<void>;
    onRunAll?: () => Promise<void>;
    onRunWithFilters?: () => Promise<void>;
    onDidChangeWorkbookTree?: vscode.Event<unknown>;
}

export function openVbaTestsPanel(
    context: vscode.ExtensionContext,
    filePath: string,
    options: VbaTestsPanelOptions,
): vscode.WebviewPanel {
    let disposed = false;
    let refreshTimer: ReturnType<typeof setTimeout> | undefined;
    let refreshVersion = 0;
    const panel = vscode.window.createWebviewPanel(
        'xlideVbaTests',
        `XLIDE Tests: ${path.basename(filePath)}`,
        vscode.ViewColumn.Beside,
        {
            enableScripts: true,
            retainContextWhenHidden: true,
        },
    );

    const renderPanel = async (): Promise<void> => {
        const model = await options.getModel();
        if (disposed) {
            return;
        }
        panel.title = `XLIDE Tests: ${model.workbookName}`;
        panel.webview.html = renderVbaTestsHtml(panel.webview, model);
    };

    const refreshPanel = async (requestVersion: number): Promise<void> => {
        if (disposed || requestVersion !== refreshVersion) {
            return;
        }
        await renderPanel();
    };

    const scheduleRefresh = (): void => {
        const requestVersion = ++refreshVersion;
        if (disposed) {
            return;
        }
        if (refreshTimer) {
            clearTimeout(refreshTimer);
        }
        refreshTimer = setTimeout(() => {
            refreshTimer = undefined;
            void refreshPanel(requestVersion).catch((err) => {
                const error = err instanceof Error ? err.message : String(err);
                void panel.webview.postMessage({ type: 'error', error });
            });
        }, 250);
    };

    const runAndRefresh = async (
        operation: (() => Promise<void>) | undefined,
        missingMessage: string,
    ): Promise<void> => {
        if (!operation) {
            await panel.webview.postMessage({ type: 'error', error: missingMessage });
            return;
        }
        await operation();
        await renderPanel();
    };

    void renderPanel().catch((err) => {
        const error = err instanceof Error ? err.message : String(err);
        panel.webview.html = renderVbaTestsErrorHtml(panel.webview, path.basename(filePath), error);
    });

    const messageSub = panel.webview.onDidReceiveMessage(async (message: { type?: string }) => {
        try {
            if (message.type === 'installSupport') {
                await runAndRefresh(options.onInstallSupport, 'XLIDE test support installation is not available.');
                await panel.webview.postMessage({ type: 'refreshed' });
                return;
            }
            if (message.type === 'runAll') {
                await runAndRefresh(options.onRunAll, 'XLIDE test execution is not available.');
                return;
            }
            if (message.type === 'runWithFilters') {
                await runAndRefresh(options.onRunWithFilters, 'XLIDE filtered test execution is not available.');
            }
        } catch (err) {
            const error = err instanceof Error ? err.message : String(err);
            await panel.webview.postMessage({ type: 'error', error });
            await renderPanel().catch(() => { /* keep existing error visible */ });
        }
    });

    const treeSub = options.onDidChangeWorkbookTree?.(() => scheduleRefresh());
    const panelDisposables = [
        messageSub,
        ...(treeSub ? [treeSub] : []),
    ];
    panel.onDidDispose(() => {
        disposed = true;
        if (refreshTimer) {
            clearTimeout(refreshTimer);
            refreshTimer = undefined;
        }
        for (const sub of panelDisposables) {
            sub.dispose();
        }
    });

    context.subscriptions.push(panel);
    return panel;
}

export function renderVbaTestsHtml(
    webviewOrModel: vscode.Webview | VbaTestsPanelModel,
    maybeModel?: VbaTestsPanelModel,
): string {
    const model = maybeModel ?? webviewOrModel as VbaTestsPanelModel;
    const webview = maybeModel ? webviewOrModel as vscode.Webview : undefined;
    const nonce = randomNonce();
    const cspSource = webview?.cspSource ?? 'vscode-resource:';
    const runEnabled = model.support.canRun && model.runtime.canRun;
    const runDisabled = runEnabled ? '' : 'disabled';
    const installDisabled = model.support.canInstall ? '' : 'disabled';
    const installTitle = model.support.canInstall
        ? `${model.support.actionLabel} ${model.workbookName}`
        : model.support.description;
    const runHelp = !model.support.canRun
        ? 'Install the bundled XlideAssert.bas support module before running workbook tests.'
        : !model.runtime.canRun
            ? model.runtime.description
            : '';

    return /* html */`<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src ${cspSource} 'unsafe-inline'; script-src 'nonce-${nonce}';">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>XLIDE Tests</title>
    <style>
        body {
            margin: 0;
            color: var(--vscode-foreground);
            background: var(--vscode-editor-background);
            font-family: var(--vscode-font-family);
            font-size: var(--vscode-font-size);
        }
        .shell {
            max-width: 1040px;
            padding: 22px;
        }
        header {
            display: flex;
            align-items: flex-start;
            justify-content: space-between;
            gap: 16px;
            border-bottom: 1px solid var(--vscode-panel-border);
            padding-bottom: 16px;
        }
        h1 {
            margin: 0 0 4px;
            font-size: 20px;
            font-weight: 700;
        }
        h2 {
            margin: 0;
            font-size: 15px;
            font-weight: 700;
        }
        .subtle {
            color: var(--vscode-descriptionForeground);
            line-height: 1.45;
        }
        button {
            min-height: 32px;
            border: 1px solid var(--vscode-button-border, transparent);
            border-radius: 4px;
            padding: 5px 12px;
            color: var(--vscode-button-foreground);
            background: var(--vscode-button-background);
            font: inherit;
            font-weight: 600;
            cursor: pointer;
        }
        button:hover:not(:disabled) {
            background: var(--vscode-button-hoverBackground);
        }
        button.secondary {
            color: var(--vscode-button-secondaryForeground);
            background: var(--vscode-button-secondaryBackground);
        }
        button.secondary:hover:not(:disabled) {
            background: var(--vscode-button-secondaryHoverBackground);
        }
        button:disabled {
            cursor: not-allowed;
            opacity: 0.55;
        }
        main {
            display: grid;
            gap: 16px;
            padding-top: 16px;
        }
        section {
            border: 1px solid var(--vscode-panel-border);
            border-radius: 6px;
            background: var(--vscode-sideBar-background);
        }
        .sectionHeader {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 12px;
            padding: 13px 15px;
            border-bottom: 1px solid var(--vscode-panel-border);
        }
        .sectionBody {
            padding: 15px;
        }
        .statusRow {
            display: grid;
            grid-template-columns: auto 1fr;
            gap: 12px;
            align-items: start;
        }
        .statusDot {
            width: 11px;
            height: 11px;
            margin-top: 4px;
            border-radius: 50%;
            background: var(--vscode-descriptionForeground);
        }
        .status-installed .statusDot {
            background: var(--vscode-testing-iconPassed, #73c991);
        }
        .status-missing .statusDot,
        .status-outdated .statusDot,
        .status-unknown .statusDot {
            background: var(--vscode-testing-iconQueued, #cca700);
        }
        .status-blocked .statusDot {
            background: var(--vscode-testing-iconErrored, #f14c4c);
        }
        .statusTitle {
            margin-bottom: 3px;
            font-weight: 700;
        }
        .helpText {
            max-width: 760px;
            color: var(--vscode-descriptionForeground);
            line-height: 1.5;
        }
        .runGrid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(210px, 1fr));
            gap: 10px;
        }
        .runButton {
            width: 100%;
            min-height: 44px;
            text-align: left;
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
        @media (max-width: 620px) {
            .shell {
                padding: 16px;
            }
            header,
            .sectionHeader {
                display: grid;
            }
        }
    </style>
</head>
<body>
    <div class="shell">
        <header>
            <div>
                <h1>XLIDE Unit Tests</h1>
                <div class="subtle">${escapeHtml(model.workbookName)}</div>
            </div>
        </header>
        <main>
            <section class="status-${escapeAttr(model.support.state)}">
                <div class="sectionHeader">
                    <h2>Test Support</h2>
                    <button type="button" data-action="installSupport" title="${escapeAttr(installTitle)}" ${installDisabled}>${escapeHtml(model.support.actionLabel)}</button>
                </div>
                <div class="sectionBody">
                    <div class="statusRow">
                        <span class="statusDot"></span>
                        <div>
                            <div class="statusTitle">${escapeHtml(model.support.title)}</div>
                            <div class="helpText">${escapeHtml(model.support.description)}</div>
                        </div>
                    </div>
                </div>
            </section>
            <section class="status-${escapeAttr(model.runtime.state)}">
                <div class="sectionHeader">
                    <h2>Runtime</h2>
                </div>
                <div class="sectionBody">
                    <div class="statusRow">
                        <span class="statusDot"></span>
                        <div>
                            <div class="statusTitle">${escapeHtml(model.runtime.title)}</div>
                            <div class="helpText">${escapeHtml(model.runtime.description)}</div>
                        </div>
                    </div>
                </div>
            </section>
            <section>
                <div class="sectionHeader">
                    <h2>Run</h2>
                </div>
                <div class="sectionBody">
                    <div class="runGrid">
                        <button class="runButton" type="button" data-action="runAll" ${runDisabled}>Run All Tests</button>
                        <button class="runButton secondary" type="button" data-action="runWithFilters" ${runDisabled}>Run With Filters</button>
                    </div>
                    ${runHelp ? `<p class="helpText">${escapeHtml(runHelp)}</p>` : ''}
                </div>
            </section>
        </main>
    </div>
    <div class="toast" id="toast"></div>
    <script nonce="${nonce}">
        const vscode = acquireVsCodeApi();
        const toast = document.getElementById('toast');
        let toastTimer;

        function showToast(message) {
            toast.textContent = message;
            toast.classList.add('visible');
            clearTimeout(toastTimer);
            toastTimer = setTimeout(() => toast.classList.remove('visible'), 2600);
        }

        document.addEventListener('click', (event) => {
            const button = event.target.closest?.('[data-action]');
            if (!button || button.disabled) {
                return;
            }
            vscode.postMessage({ type: button.dataset.action });
        });

        window.addEventListener('message', (event) => {
            if (event.data?.type === 'error') {
                showToast(event.data.error || 'XLIDE test action failed');
            } else if (event.data?.type === 'refreshed') {
                showToast('Test support refreshed');
            }
        });
    </script>
</body>
</html>`;
}

function renderVbaTestsErrorHtml(webview: vscode.Webview, workbookName: string, error: string): string {
    return /* html */`<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src ${webview.cspSource} 'unsafe-inline';">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>XLIDE Tests Error</title>
    <style>
        body {
            margin: 0;
            padding: 24px;
            color: var(--vscode-foreground);
            background: var(--vscode-editor-background);
            font-family: var(--vscode-font-family);
        }
        h1 {
            margin: 0 0 6px;
            font-size: 18px;
        }
        .subtle {
            color: var(--vscode-descriptionForeground);
            margin-bottom: 18px;
        }
        .error {
            border: 1px solid var(--vscode-inputValidation-errorBorder);
            border-radius: 4px;
            padding: 12px;
            white-space: pre-wrap;
        }
    </style>
</head>
<body>
    <main>
        <h1>XLIDE Tests Could Not Load</h1>
        <div class="subtle">${escapeHtml(workbookName)}</div>
        <div class="error">${escapeHtml(error)}</div>
    </main>
</body>
</html>`;
}

function randomNonce(): string {
    const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    let nonce = '';
    for (let i = 0; i < 32; i++) {
        nonce += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    return nonce;
}

function escapeHtml(value: unknown): string {
    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function escapeAttr(value: unknown): string {
    return escapeHtml(value);
}
