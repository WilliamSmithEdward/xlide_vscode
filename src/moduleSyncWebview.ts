import * as vscode from 'vscode';
import type { ModuleSyncPlan } from './moduleSyncPlan';

export interface ModuleSyncApplyResult {
    summary: string;
    changed: number;
    skipped: number;
    removed?: number;
    failed: number;
}

export function openModuleSyncPreview(
    context: vscode.ExtensionContext,
    plan: ModuleSyncPlan,
    onApply: (selectedIds: readonly string[]) => Promise<ModuleSyncApplyResult>,
): Promise<ModuleSyncApplyResult | undefined> {
    return new Promise((resolve) => {
        const panel = vscode.window.createWebviewPanel(
            'xlide.moduleSyncPreview',
            plan.direction === 'export' ? 'XLIDE Export Preview' : 'XLIDE Import Preview',
            vscode.ViewColumn.One,
            {
                enableScripts: true,
                retainContextWhenHidden: true,
                localResourceRoots: [context.extensionUri],
            },
        );
        let resolved = false;
        const done = (result: ModuleSyncApplyResult | undefined): void => {
            if (!resolved) {
                resolved = true;
                resolve(result);
            }
        };

        panel.webview.html = renderModuleSyncHtml(panel.webview, plan);
        const messageSub = panel.webview.onDidReceiveMessage(async (message: {
            type?: string;
            selectedIds?: string[];
        }) => {
            if (message.type === 'cancel') {
                done(undefined);
                panel.dispose();
                return;
            }
            if (message.type !== 'apply') {
                return;
            }
            const selectedIds = message.selectedIds ?? [];
            await panel.webview.postMessage({ type: 'applying' });
            try {
                const result = await onApply(selectedIds);
                await panel.webview.postMessage({ type: 'applied', result });
                done(result);
            } catch (err) {
                const error = err instanceof Error ? err.message : String(err);
                await panel.webview.postMessage({ type: 'error', error });
            }
        });
        panel.onDidDispose(() => {
            messageSub.dispose();
            done(undefined);
        });
    });
}

function renderModuleSyncHtml(webview: vscode.Webview, plan: ModuleSyncPlan): string {
    const nonce = getNonce();
    const data = JSON.stringify(plan).replace(/</g, '\\u003c');
    return /* html */`<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src ${webview.cspSource} 'unsafe-inline'; script-src 'nonce-${nonce}';">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${escapeHtml(plan.title)}</title>
    <style>
        :root {
            --border: color-mix(in srgb, var(--vscode-editor-foreground) 18%, transparent);
            --muted: var(--vscode-descriptionForeground);
            --row: color-mix(in srgb, var(--vscode-editor-background) 90%, var(--vscode-editor-foreground) 10%);
            --changed: color-mix(in srgb, var(--vscode-gitDecoration-modifiedResourceForeground) 16%, transparent);
            --added: color-mix(in srgb, var(--vscode-gitDecoration-addedResourceForeground) 18%, transparent);
            --removed: color-mix(in srgb, var(--vscode-gitDecoration-deletedResourceForeground) 16%, transparent);
            --warn: color-mix(in srgb, var(--vscode-editorWarning-foreground) 18%, transparent);
        }
        * { box-sizing: border-box; }
        body {
            margin: 0;
            color: var(--vscode-editor-foreground);
            background: var(--vscode-editor-background);
            font: var(--vscode-font-size) var(--vscode-font-family);
            height: 100vh;
            overflow: hidden;
        }
        .shell { display: grid; grid-template-rows: auto 1fr auto; height: 100vh; }
        header {
            padding: 12px 16px;
            border-bottom: 1px solid var(--border);
            display: grid;
            grid-template-columns: 1fr auto;
            gap: 12px;
            align-items: center;
        }
        h1 {
            margin: 0 0 4px;
            font-size: 16px;
            font-weight: 600;
        }
        .sub { color: var(--muted); font-size: 12px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
        .actions { display: flex; gap: 8px; align-items: center; }
        button {
            color: var(--vscode-button-foreground);
            background: var(--vscode-button-background);
            border: 0;
            padding: 5px 10px;
            min-height: 28px;
            border-radius: 2px;
            cursor: pointer;
        }
        button.secondary {
            color: var(--vscode-button-secondaryForeground);
            background: var(--vscode-button-secondaryBackground);
        }
        button:disabled { opacity: .55; cursor: default; }
        main { display: grid; grid-template-columns: minmax(280px, 34%) 1fr; min-height: 0; }
        aside { border-right: 1px solid var(--border); min-height: 0; display: grid; grid-template-rows: auto auto 1fr; }
        .toolbar {
            display: flex;
            gap: 8px;
            padding: 8px;
            border-bottom: 1px solid var(--border);
        }
        .toolbar button { flex: 1; }
        .warnings {
            padding: 8px 12px;
            display: none;
            background: var(--warn);
            border-bottom: 1px solid var(--border);
            color: var(--vscode-editorWarning-foreground);
            line-height: 1.35;
        }
        .warnings.visible { display: block; }
        .list { overflow: auto; }
        .item {
            width: 100%;
            display: grid;
            grid-template-columns: auto 1fr auto;
            gap: 8px;
            align-items: center;
            padding: 8px 10px;
            border-bottom: 1px solid var(--border);
            cursor: pointer;
        }
        .item:hover, .item.active { background: var(--row); }
        .item input { margin: 0; }
        .name { font-weight: 600; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
        .meta { color: var(--muted); font-size: 12px; margin-top: 2px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
        .badge {
            font-size: 11px;
            padding: 2px 6px;
            border-radius: 999px;
            border: 1px solid var(--border);
            white-space: nowrap;
        }
        .badge.skip { color: var(--vscode-editorWarning-foreground); background: var(--warn); }
        .badge.same { color: var(--muted); }
        section { min-height: 0; display: grid; grid-template-rows: auto 1fr; }
        .diff-head {
            padding: 10px 12px;
            border-bottom: 1px solid var(--border);
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 8px;
        }
        .diff-title { font-weight: 600; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
        .diff { overflow: auto; font-family: var(--vscode-editor-font-family); font-size: var(--vscode-editor-font-size); }
        .line {
            display: grid;
            grid-template-columns: 56px minmax(0, 1fr) 56px minmax(0, 1fr);
            min-height: 20px;
            border-bottom: 1px solid color-mix(in srgb, var(--border) 40%, transparent);
        }
        .ln {
            color: var(--muted);
            text-align: right;
            padding: 2px 8px;
            user-select: none;
            border-right: 1px solid var(--border);
        }
        pre {
            margin: 0;
            padding: 2px 8px;
            white-space: pre-wrap;
            overflow-wrap: anywhere;
            min-width: 0;
        }
        .line.changed pre { background: var(--changed); }
        .line.added pre.right { background: var(--added); }
        .line.removed pre.left { background: var(--removed); }
        footer {
            min-height: 34px;
            padding: 8px 12px;
            border-top: 1px solid var(--border);
            color: var(--muted);
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 12px;
        }
        .result { color: var(--vscode-editor-foreground); }
    </style>
</head>
<body>
    <div class="shell">
        <header>
            <div>
                <h1 id="title"></h1>
                <div class="sub" id="subtitle"></div>
            </div>
            <div class="actions">
                <button class="secondary" id="cancel">Cancel</button>
                <button id="apply">Apply Selected</button>
            </div>
        </header>
        <main>
            <aside>
                <div class="toolbar">
                    <button class="secondary" id="selectChanged">Select Changed</button>
                    <button class="secondary" id="clear">Clear</button>
                </div>
                <div class="warnings" id="warnings"></div>
                <div class="list" id="list"></div>
            </aside>
            <section>
                <div class="diff-head">
                    <div class="diff-title" id="leftTitle"></div>
                    <div class="diff-title" id="rightTitle"></div>
                </div>
                <div class="diff" id="diff"></div>
            </section>
        </main>
        <footer>
            <span id="counts"></span>
            <span class="result" id="result"></span>
        </footer>
    </div>
    <script nonce="${nonce}">
        const vscode = acquireVsCodeApi();
        const plan = ${data};
        const selected = new Set(plan.items.filter(item => item.checked && item.selectable).map(item => item.id));
        let activeId = plan.items[0]?.id;
        let applying = false;
        let applied = false;

        const el = id => document.getElementById(id);
        el('title').textContent = plan.title;
        el('subtitle').textContent = \`\${plan.workbookPath} <-> \${plan.folderPath}\${plan.exportMode ? '  [' + plan.exportMode + ']' : ''}\`;
        if (plan.warnings.length) {
            el('warnings').classList.add('visible');
            el('warnings').textContent = plan.warnings.join('\\n');
        }

        function renderList() {
            const list = el('list');
            list.innerHTML = '';
            for (const item of plan.items) {
                const row = document.createElement('div');
                row.className = 'item' + (item.id === activeId ? ' active' : '');
                row.dataset.id = item.id;
                const checkbox = document.createElement('input');
                checkbox.type = 'checkbox';
                checkbox.checked = selected.has(item.id);
                checkbox.disabled = !item.selectable;
                checkbox.addEventListener('click', event => {
                    event.stopPropagation();
                    if (checkbox.checked) selected.add(item.id);
                    else selected.delete(item.id);
                    renderCounts();
                });
                const text = document.createElement('div');
                const name = document.createElement('div');
                name.className = 'name';
                name.textContent = item.moduleName;
                const meta = document.createElement('div');
                meta.className = 'meta';
                meta.textContent = [item.relativeName, item.moduleType, item.warning].filter(Boolean).join(' | ');
                text.append(name, meta);
                const badge = document.createElement('span');
                badge.className = 'badge' + (item.status.startsWith('skipping') ? ' skip' : item.status === 'unchanged' ? ' same' : '');
                badge.textContent = item.detail || item.status;
                row.append(checkbox, text, badge);
                row.addEventListener('click', () => {
                    activeId = item.id;
                    renderList();
                    renderDiff();
                });
                list.append(row);
            }
            renderCounts();
        }

        function renderDiff() {
            const item = plan.items.find(candidate => candidate.id === activeId) || plan.items[0];
            if (!item) return;
            el('leftTitle').textContent = item.leftTitle;
            el('rightTitle').textContent = item.rightTitle;
            const diff = el('diff');
            diff.innerHTML = '';
            for (const line of item.diff) {
                const row = document.createElement('div');
                row.className = 'line ' + line.kind;
                const leftNo = document.createElement('div');
                leftNo.className = 'ln';
                leftNo.textContent = line.leftNumber || '';
                const left = document.createElement('pre');
                left.className = 'left';
                left.textContent = line.left;
                const rightNo = document.createElement('div');
                rightNo.className = 'ln';
                rightNo.textContent = line.rightNumber || '';
                const right = document.createElement('pre');
                right.className = 'right';
                right.textContent = line.right;
                row.append(leftNo, left, rightNo, right);
                diff.append(row);
            }
        }

        function renderCounts() {
            const selectedItems = plan.items.filter(item => selected.has(item.id));
            const unsupported = selectedItems.filter(item => item.unsupportedDirectCreation).length;
            el('counts').textContent = \`\${selectedItems.length} selected\${unsupported ? ' | ' + unsupported + ' will show skipping import warning' : ''}\`;
            el('apply').disabled = applying || applied || selectedItems.length === 0;
        }

        el('selectChanged').addEventListener('click', () => {
            selected.clear();
            for (const item of plan.items) {
                if (item.selectable && item.status !== 'unchanged' && !item.status.startsWith('skipping')) selected.add(item.id);
            }
            renderList();
        });
        el('clear').addEventListener('click', () => {
            selected.clear();
            renderList();
        });
        el('cancel').addEventListener('click', () => vscode.postMessage({ type: 'cancel' }));
        el('apply').addEventListener('click', () => {
            if (applied) return;
            applying = true;
            el('result').textContent = 'Applying...';
            renderCounts();
            vscode.postMessage({ type: 'apply', selectedIds: Array.from(selected) });
        });

        window.addEventListener('message', event => {
            const message = event.data;
            if (message.type === 'applying') {
                applying = true;
                el('result').textContent = 'Applying...';
                renderCounts();
            } else if (message.type === 'applied') {
                applying = false;
                applied = true;
                el('result').textContent = message.result.summary;
                el('apply').textContent = 'Applied';
                renderCounts();
            } else if (message.type === 'error') {
                applying = false;
                el('result').textContent = message.error;
                renderCounts();
            }
        });

        renderList();
        renderDiff();
    </script>
</body>
</html>`;
}

function getNonce(): string {
    const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    let text = '';
    for (let i = 0; i < 32; i++) {
        text += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    return text;
}

function escapeHtml(value: string): string {
    return value
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}
