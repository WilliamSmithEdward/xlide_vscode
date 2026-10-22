import * as vscode from 'vscode';
import type { CommandDeps } from './commands/shared';
import type { XlideNode } from './projectExplorer';
import { encodeModuleUri, decodeModuleUri, XLIDE_SCHEME } from './xlideFileSystem';
import { moduleIdentityKey, projectIdentityKey } from './projectIdentity';
import { workspaceUriFor } from './util/workspaceUris';
import { errorMessage } from './util/errors';
import { readExtensionTextAsset } from './extensionAssets';
import { randomNonce, scriptJson } from './webview/html';
import { moduleSearchChangePreview, moduleSearchPreviewer, moduleSearchReplacement, validateModuleSearchMatches,
    type ModuleSearchMatch, type ModuleSearchQuery } from './moduleSearchModel';

const MAX_MATCHES = 5000;
const MAX_MODULE_LENGTH = 2_000_000;
const MAX_SEARCH_LENGTH = 20_000_000;

interface Hit {
    id: number;
    uri: vscode.Uri;
    project: string;
    module: string;
    source: string;
    match: ModuleSearchMatch;
}

function documentKey(uri: vscode.Uri): string {
    if (uri.scheme === XLIDE_SCHEME) {
        const { projectPath, moduleName, face } = decodeModuleUri(uri);
        return JSON.stringify([projectIdentityKey(projectPath), moduleIdentityKey(moduleName), face ?? 'code']);
    }
    return projectIdentityKey(uri.toString());
}
export interface PlannedModule { doc: vscode.TextDocument; source: string; edits: Array<{ hit: Hit; text: string }> }

/** Apply only after preview/confirmation; VS Code owns undo and normal save behavior. */
export async function applyModuleSearchPlan(plan: readonly PlannedModule[]): Promise<void> {
    const edit = new vscode.WorkspaceEdit();
    for (const module of plan) {
        if (module.doc.isClosed || module.doc.getText() !== module.source) {
            throw new Error('A module changed after preview. Search again; nothing was applied.');
        }
        for (const { hit, text } of module.edits) {
            edit.replace(module.doc.uri, new vscode.Range(module.doc.positionAt(hit.match.start), module.doc.positionAt(hit.match.end)), text);
        }
    }
    if (!await vscode.workspace.applyEdit(edit)) { throw new Error('VS Code could not apply the replacements.'); }
}

/** All filesystem identities come from the tree/engine, never from webview messages. */
export class ModuleSearchPanel implements vscode.Disposable {
    private readonly panel: vscode.WebviewPanel;
    private readonly disposables: vscode.Disposable[] = [];
    private generation = 0;
    private previewGeneration = 0;
    private requestId = 0;
    private pending?: { id: number; resolve(value: unknown): void; reject(error: Error): void; timer: ReturnType<typeof setTimeout> };
    private hits: Hit[] = [];
    private plan: PlannedModule[] = [];
    private query?: ModuleSearchQuery;
    private incomplete = false;
    private busy = false;
    private disposed = false;

    constructor(private readonly deps: CommandDeps, private readonly origin?: XlideNode) {
        this.panel = vscode.window.createWebviewPanel('xlide.moduleSearch', 'XLIDE Find / Replace',
            vscode.ViewColumn.Beside, { enableScripts: true, localResourceRoots: [], retainContextWhenHidden: true });
        const nonce = randomNonce();
        const worker = readExtensionTextAsset('assets/webview/moduleSearchWorker.js');
        const ui = readExtensionTextAsset('assets/webview/moduleSearch.js');
        this.panel.webview.html = readExtensionTextAsset('assets/webview/moduleSearch.html')
            .replaceAll('{{nonce}}', nonce)
            .replace('{{worker}}', () => scriptJson(worker))
            .replace('{{script}}', () => ui);
        this.disposables.push(this.panel.webview.onDidReceiveMessage(message => {
            void this.receive(message).catch(error => {
                if (!this.disposed) { this.send({ type: 'error', text: errorMessage(error) }); }
            });
        }), this.panel.onDidDispose(() => this.dispose()));
    }

    private send(message: object): void { if (!this.disposed) { void this.panel.webview.postMessage(message); } }

    private invalidate(): void {
        this.generation++;
        this.previewGeneration++;
        this.hits = [];
        this.plan = [];
        this.query = undefined;
        this.busy = false;
        const pending = this.pending;
        this.pending = undefined;
        if (pending) { clearTimeout(pending.timer); pending.reject(new Error('Search cancelled.')); }
        this.send({ type: 'resetWorker' });
    }

    dispose(): void {
        if (this.disposed) { return; }
        this.invalidate();
        this.disposed = true;
        for (const disposable of this.disposables.splice(0)) { disposable.dispose(); }
        this.panel.dispose();
    }

    private async receive(message: unknown): Promise<void> {
        if (!message || typeof message !== 'object') { return; }
        const m = message as Record<string, unknown>;
        if (m.type === 'matched' && m.id === this.pending?.id) {
            const pending = this.pending!;
            this.pending = undefined;
            clearTimeout(pending.timer);
            if (typeof m.error === 'string') { pending.reject(new Error(m.error)); }
            else { pending.resolve(m.result); }
            return;
        }
        if (m.type === 'ready') {
            this.send({ type: 'scope', project: this.origin?.filePath ?? '', module: this.origin?.moduleName ?? '',
                projectAvailable: !!this.origin?.filePath, moduleAvailable: !!this.origin?.moduleName,
                selected: this.origin?.moduleName ? 'module' : this.origin?.filePath ? 'project' : 'all' });
        } else if (m.type === 'invalidate' || m.type === 'stop') {
            this.invalidate();
            this.send({ type: 'stopped' });
        } else if (m.type === 'replacementChanged') {
            this.plan = [];
            this.previewGeneration++;
        } else if (m.type === 'search') {
            await this.search(m);
        } else if (!this.busy && m.type === 'open' && Number.isInteger(m.id)) {
            const hit = this.hits.find(hit => hit.id === m.id);
            if (!hit) { return; }
            const doc = await vscode.workspace.openTextDocument(hit.uri);
            if (doc.getText() !== hit.source) { throw new Error('The module changed. Search again before navigating.'); }
            const range = new vscode.Range(doc.positionAt(hit.match.start), doc.positionAt(hit.match.end));
            await vscode.window.showTextDocument(doc, { selection: range, preview: true });
        } else if (!this.busy && m.type === 'preview' && typeof m.replacement === 'string') {
            if (m.replacement.length > 100_000) { throw new Error('Replacement text is too long.'); }
            await this.preview(m.ids, m.replacement, m.all === true);
        } else if (!this.busy && m.type === 'apply') {
            await this.apply();
        }
    }

    private match(text: string, query: ModuleSearchQuery, limit: number): Promise<unknown> {
        return new Promise((resolve, reject) => {
            const id = ++this.requestId;
            const timer = setTimeout(() => {
                if (this.pending?.id !== id) { return; }
                this.pending = undefined;
                this.send({ type: 'resetWorker' });
                reject(new Error('Matching exceeded two seconds. Simplify the pattern or narrow the scope.'));
            }, 2000);
            this.pending = { id, resolve, reject, timer };
            this.send({ type: 'match', id, text, query, limit });
        });
    }

    private async search(message: Record<string, unknown>): Promise<void> {
        this.invalidate();
        const generation = this.generation;
        const current = () => generation === this.generation && !this.disposed;
        if (typeof message.pattern !== 'string' || !message.pattern || message.pattern.length > 1024) {
            throw new Error('Enter a pattern of 1–1024 characters.');
        }
        const query: ModuleSearchQuery = { pattern: message.pattern, regex: message.regex === true,
            matchCase: message.matchCase === true, wholeWord: message.wholeWord === true };
        const scope = message.scope;
        if (!['all', 'project', 'module'].includes(String(scope))) { throw new Error('Choose a search scope.'); }
        const projects = (await this.deps.explorer.getChildren()).filter(node => node.kind === 'project');
        if (!current()) { return; }
        const targets = scope === 'all' ? projects : projects.filter(p => p.filePath === this.origin?.filePath);
        if (!targets.length || (scope === 'module' && !this.origin?.moduleName)) { throw new Error('No matching project or module in the XLIDE tree.'); }
        this.query = query;
        this.busy = true;
        this.incomplete = false;
        this.send({ type: 'busy' });
        const errors: string[] = [];
        let length = 0;
        try {
            for (const project of targets) {
                if (!current() || this.hits.length >= MAX_MATCHES) { break; }
                try {
                    const modules = await this.deps.bridge.call<Array<{ name: string; filePath?: string }>>('listModules', { path: project.filePath });
                    for (const module of modules) {
                        if (!current()) { return; }
                        if (scope === 'module' && module.name !== this.origin?.moduleName) { continue; }
                        let uri = module.filePath ? workspaceUriFor(module.filePath) : encodeModuleUri(project.filePath, module.name);
                        const key = documentKey(uri);
                        const open = vscode.workspace.textDocuments.find(doc => {
                            try { return documentKey(doc.uri) === key; } catch { return false; }
                        });
                        if (open) { uri = open.uri; }
                        const source = open ? open.getText() : Buffer.from(module.filePath
                            ? await vscode.workspace.fs.readFile(uri) : await this.deps.fsProvider.readFile(uri)).toString('utf8');
                        if (!current()) { return; }
                        length += source.length;
                        if (source.length > MAX_MODULE_LENGTH || length > MAX_SEARCH_LENGTH) {
                            throw new Error('Search size limit reached. Narrow the scope.');
                        }
                        const remaining = MAX_MATCHES - this.hits.length;
                        if (remaining <= 0) { this.incomplete = true; break; }
                        const response = await this.match(source, query, remaining) as { matches?: unknown; limitHit?: unknown };
                        if (!current()) { return; }
                        const matches = validateModuleSearchMatches(source, response?.matches, remaining);
                        if (response.limitHit === true) { this.incomplete = true; }
                        if (!matches.length) { continue; }
                        const preview = moduleSearchPreviewer(source);
                        const rows = matches.map(match => {
                            const hit: Hit = { id: this.hits.length, uri, project: project.filePath, module: module.name, source, match };
                            this.hits.push(hit);
                            return { id: hit.id, project: project.label, projectPath: project.filePath, module: module.name, ...preview(match) };
                        });
                        this.send({ type: 'results', rows });
                    }
                } catch (error) {
                    if (!current()) { return; }
                    this.incomplete = true;
                    errors.push(`${project.label}: ${errorMessage(error)}`);
                }
            }
            if (this.hits.length >= MAX_MATCHES) { this.incomplete = true; }
            if (current()) { this.send({ type: 'done', count: this.hits.length, incomplete: this.incomplete, errors }); }
        } finally { if (current()) { this.busy = false; } }
    }

    private async preview(ids: unknown, replacement: string, all: boolean): Promise<void> {
        if (!this.query) { throw new Error('Search before preparing replacements.'); }
        if (all && this.incomplete) { throw new Error('Search was incomplete. Narrow the scope before replacing all.'); }
        if (!all && (!Array.isArray(ids) || ids.length > MAX_MATCHES || !ids.every(Number.isInteger))) {
            throw new Error('Select matches to replace.');
        }
        const wanted = new Set(all ? this.hits.map(h => h.id) : ids as number[]);
        const hits = this.hits.filter(hit => wanted.has(hit.id));
        if (!hits.length || hits.length !== wanted.size) { throw new Error('Select current search matches.'); }
        const generation = this.generation;
        const previewGeneration = ++this.previewGeneration;
        this.busy = true;
        this.plan = [];
        try {
            const plan = new Map<string, PlannedModule>();
            const rows = [];
            let size = 0;
            for (const hit of hits) {
                const key = hit.uri.toString();
                let module = plan.get(key);
                if (!module) {
                    await this.deps.bridge.call('ensureProjectAccess', { path: hit.project });
                    const doc = await vscode.workspace.openTextDocument(hit.uri);
                    if (doc.getText() !== hit.source) { throw new Error(`${hit.module} changed. Search again.`); }
                    module = { doc, source: hit.source, edits: [] };
                    plan.set(key, module);
                }
                if (generation !== this.generation || previewGeneration !== this.previewGeneration || !this.query) { return; }
                const text = moduleSearchReplacement(hit.source, hit.match, replacement, this.query.regex, MAX_SEARCH_LENGTH - size);
                size += text.length;
                if (size > MAX_SEARCH_LENGTH) { throw new Error('Replacement size limit reached. Narrow the scope.'); }
                module.edits.push({ hit, text });
                rows.push({ id: hit.id, ...moduleSearchChangePreview(hit.source, hit.match, text) });
            }
            if (generation !== this.generation || previewGeneration !== this.previewGeneration) { return; }
            this.plan = [...plan.values()];
            this.send({ type: 'preview', rows, count: hits.length, modules: this.plan.length });
        } finally {
            if (generation === this.generation) { this.busy = false; this.send({ type: 'idle', hasPreview: this.plan.length > 0 }); }
        }
    }

    private async apply(): Promise<void> {
        if (!this.plan.length) { throw new Error('Preview replacements first.'); }
        const plan = this.plan;
        const generation = this.generation;
        this.busy = true;
        try {
            const count = plan.reduce((sum, module) => sum + module.edits.length, 0);
            const answer = await vscode.window.showWarningMessage(
                `Apply ${count} replacement(s) in ${plan.length} module(s)? Normal Save / Auto Save settings apply. Use Undo to revert.`,
                { modal: true }, 'Apply replacements');
            if (answer !== 'Apply replacements' || generation !== this.generation || this.plan !== plan) { return; }
            await applyModuleSearchPlan(plan);
            this.invalidate();
            this.send({ type: 'applied', count });
        } finally {
            if (generation === this.generation) { this.busy = false; this.send({ type: 'idle', hasPreview: this.plan.length > 0 }); }
        }
    }
}
