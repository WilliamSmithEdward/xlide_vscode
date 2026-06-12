import { escapeHtml, randomNonce } from './html';

/** Charset/CSP/viewport/title head block shared by all XLIDE webviews. */
export function webviewHeadHtml(nonce: string, title: string, options: { allowScripts?: boolean } = {}): string {
    const scriptSrc = options.allowScripts === false ? '' : ` script-src 'nonce-${nonce}';`;
    return `<meta charset="UTF-8">
    <meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'nonce-${nonce}';${scriptSrc}">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${escapeHtml(title)}</title>`;
}

/** Script-free error page shared by the feature panels. */
export function renderWebviewErrorPageHtml(options: {
    title: string;
    heading: string;
    subtitle: string;
    error: string;
    help?: string;
}): string {
    const nonce = randomNonce();
    return /* html */`<!DOCTYPE html>
<html lang="en">
<head>
    ${webviewHeadHtml(nonce, options.title, { allowScripts: false })}
    <style nonce="${nonce}">
        body {
            margin: 0;
            padding: 24px;
            color: var(--vscode-foreground);
            background: var(--vscode-editor-background);
            font-family: var(--vscode-font-family);
        }
        main {
            max-width: 900px;
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
            background: color-mix(in srgb, var(--vscode-inputValidation-errorBackground) 40%, transparent);
            padding: 12px;
            border-radius: 4px;
            line-height: 1.45;
            white-space: pre-wrap;
        }
        .help {
            margin-top: 12px;
            color: var(--vscode-descriptionForeground);
        }
    </style>
</head>
<body>
    <main>
        <h1>${escapeHtml(options.heading)}</h1>
        <div class="subtle">${escapeHtml(options.subtitle)}</div>
        <div class="error">${escapeHtml(options.error)}</div>
        ${options.help ? `<div class="help">${escapeHtml(options.help)}</div>` : ''}
    </main>
</body>
</html>`;
}

export function statHtml(value: string | number, label: string): string {
    return `<div class="stat"><strong>${escapeHtml(String(value))}</strong><span>${escapeHtml(label)}</span></div>`;
}

export const WEBVIEW_TOAST_HTML = '<div class="toast" id="toast"></div>';

export const WEBVIEW_TOAST_CSS = `.toast {
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
        }`;

/** In-webview toast helper; expects the shared toast markup and CSS. */
export const WEBVIEW_TOAST_SCRIPT = `const toast = document.getElementById('toast');
        let toastTimer;
        function showToast(message) {
            toast.textContent = message;
            toast.classList.add('visible');
            clearTimeout(toastTimer);
            toastTimer = setTimeout(() => toast.classList.remove('visible'), 2600);
        }`;
