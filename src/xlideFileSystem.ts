import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';
import { ProjectEngine } from './projectEngine';
import { errorCategoryForSupportLog, WORKBOOK_LOCKED_ERROR_RE } from './xlideCommandLog';
import { formatChangeSummary, recordXlideWriteAudit } from './xlideWriteAudit';
import { startPerformanceTrace } from './performanceTrace';
import { errorMessage } from './util/errors';
import { runWriteWithExcelCoordination } from './excelWorkbookCoordinator';
import { noteModuleWrite } from './vbaRenameHistory';
// Function-level cycle with xlideAgentDiff (it imports URI/identity helpers
// from this module); neither side touches the other at module-eval time.
import { trackModuleWriteForAgentReview } from './xlideAgentDiff';
import { containerAppNameForPath, MACRO_CONTAINER_EXTENSION_PATTERN } from './macroContainerUi';
import { projectIdentityKey } from './projectIdentity';

export const XLIDE_SCHEME = 'xlide-vba';

const MODULE_URI_RE = new RegExp(
    `^(.*\\.(?:${MACRO_CONTAINER_EXTENSION_PATTERN}))/([^/]+)\\.(bas|form)$`,
    'i',
);
export const XLIDE_VBA_LANGUAGE_ID = 'xlide-vba';

export { moduleIdentityKey, sameProjectPath, projectIdentityKey } from './projectIdentity';

/**
 * True for any VBA document: by language id or by xlide scheme. A `.form`
 * document rides the same scheme but is the form's MARKUP face, not VBA -
 * running the VBA analyzer on it painted every element as a statement
 * outside a procedure.
 */
export function isVbaDocument(document: vscode.TextDocument): boolean {
    if (document.uri.scheme === XLIDE_SCHEME) {
        return !document.uri.path.toLowerCase().endsWith('.form');
    }
    return document.languageId === 'vba'
        || document.languageId === XLIDE_VBA_LANGUAGE_ID;
}

/** True for xlide-scheme documents backed by a project on disk. */
export function isLocalXlideDocument(document: vscode.TextDocument): boolean {
    return document.uri.scheme === XLIDE_SCHEME;
}

/** The active editor when it shows a local project VBA module, else undefined. */
export function activeLocalVbaEditor(): vscode.TextEditor | undefined {
    const editor = vscode.window.activeTextEditor;
    return editor && isLocalXlideDocument(editor.document) ? editor : undefined;
}

/**
 * Tracks project paths for which the signature-dropped notice has already
 * been shown this session, so the user sees it at most once per file.
 */
const _sigWarnedPaths = new Set<string>();

/**
 * Show a one-time warning when a VBA digital signature was invalidated by a
 * save.  Safe to call on every write - suppressed after the first occurrence
 * per project path per session.
 */
export function notifySignatureDropped(filePath: string, signatureDropped: boolean): void {
    const key = projectIdentityKey(filePath);
    if (!signatureDropped || _sigWarnedPaths.has(key)) { return; }
    _sigWarnedPaths.add(key);
    void vscode.window.showWarningMessage(
        `XLIDE: "${path.basename(filePath)}" had a VBA digital signature that was invalidated by this edit. ` +
        `Re-sign the project externally to restore trust.`,
    );
}

/**
 * Heuristic: does this error string look like a Windows file-sharing violation
 * caused by Excel having the workbook open?
 */
export function isProjectLockedError(message: string): boolean {
    return WORKBOOK_LOCKED_ERROR_RE.test(message);
}

// Collapse rapid repeat lock notices for the same project into a single popup
// (e.g. a burst of operations, or a writeFile failure followed by a re-read).
const LOCKED_NOTICE_THROTTLE_MS = 2000;
const recentLockedNotices = new Map<string, number>();

export function reportProjectLocked(projectPath: string, op: 'read' | 'write'): void {
    const noticeKey = projectIdentityKey(projectPath);
    const now = Date.now();
    const last = recentLockedNotices.get(noticeKey);
    if (last !== undefined && now - last < LOCKED_NOTICE_THROTTLE_MS) {
        return;
    }
    recentLockedNotices.set(noticeKey, now);
    const name = path.basename(projectPath);
    const verb = op === 'read' ? 'open' : 'save';
    // Retry (a revert to re-read the file) only fits the READ case: on a
    // failed write it would revert whatever editor happens to be active,
    // discarding unrelated dirty edits instead of retrying anything.
    const actions = op === 'read' ? ['Retry', 'Reveal File'] : ['Reveal File'];
    void vscode.window.showWarningMessage(
        `XLIDE: Cannot ${verb} "${name}" - it appears to be open in ${containerAppNameForPath(projectPath)}. Close the file and try again.`,
        ...actions,
    ).then((choice) => {
        if (choice === 'Retry') {
            void vscode.commands.executeCommand('workbench.action.files.revert');
        } else if (choice === 'Reveal File') {
            void vscode.commands.executeCommand('revealFileInOS', vscode.Uri.file(projectPath));
        }
    });
}

/**
 * Encodes a (projectPath, moduleName) pair into a virtual URI.
 * URI form: xlide-vba:///C:/path/to/workbook.xlsm/ModuleName.bas
 */
export function encodeModuleUri(projectPath: string, moduleName: string): vscode.Uri {
    const forward = projectPath.replace(/\\/g, '/');
    const base = forward.startsWith('/') ? forward : `/${forward}`;
    // Build from a structured path rather than interpolating into a URI string.
    // vscode.Uri.from percent-encodes the path on serialization while keeping
    // uri.path literal, so project paths containing reserved characters like
    // '#' or '%' round-trip correctly. (String interpolation + Uri.parse would
    // split the path on '#'/'?' into the fragment/query and silently decode a
    // stray '%xx', breaking decode or pointing the bridge at the wrong file.)
    return vscode.Uri.from({
        scheme: XLIDE_SCHEME,
        path: `${base}/${moduleName}.bas`,
    });
}

/**
 * Encodes a form's MARKUP document: the same project path, the module name,
 * and a .form suffix so the provider routes reads and saves to the designer
 * rather than the code-behind.
 */
export function encodeFormMarkupUri(projectPath: string, moduleName: string): vscode.Uri {
    const forward = projectPath.replace(/\\/g, '/');
    const base = forward.startsWith('/') ? forward : `/${forward}`;
    return vscode.Uri.from({
        scheme: XLIDE_SCHEME,
        path: `${base}/${moduleName}.form`,
    });
}

/**
 * Decodes a virtual URI back to (projectPath, moduleName).
 */
export function decodeModuleUri(uri: vscode.Uri): { projectPath: string; moduleName: string; face?: 'code' | 'form' } {
    const p = uri.path;
    // Match the macro-container boundary in the path: any extension the
    // engine opens (.xlsm through .accdb), so modules from every container
    // open in the editor, not only Excel's.
    const match = p.match(MODULE_URI_RE);
    if (!match) {
        throw new Error(`Cannot decode xlide-vba URI: ${uri.toString()}`);
    }
    let rawPath = match[1]; // e.g. /C:/Users/.../workbook.xlsm
    const moduleName = decodeURIComponent(match[2]);
    const face = match[3]?.toLowerCase() === 'form' ? 'form' as const : 'code' as const;

    // On Windows, the leading slash before the drive letter is artificial
    if (process.platform === 'win32' && /^\/[A-Za-z]:/.test(rawPath)) {
        rawPath = rawPath.slice(1);
    }
    const projectPath = rawPath.replace(/\//g, path.sep);
    return { projectPath, moduleName, face };
}

/**
 * Virtual FileSystemProvider for the xlide-vba:// scheme.
 *
 * - readFile  -> calls the project engine's readModule
 * - writeFile -> calls the project engine's writeModule (saves the file in place)
 * - All other mutation operations are rejected.
 */
export class XlideFileSystemProvider
    implements vscode.FileSystemProvider, vscode.Disposable
{
    private _emitter = new vscode.EventEmitter<vscode.FileChangeEvent[]>();
    readonly onDidChangeFile = this._emitter.event;

    private _clock = Date.now();
    private readonly _stats = new Map<string, { ctime: number; mtime: number; size: number; projectKey?: string }>();
    /** Last known real project file mtime per project identity key. */
    private readonly _projectMtimes = new Map<string, number>();
    private readonly _disposables: vscode.Disposable[] = [];

    constructor(private readonly _bridge: ProjectEngine) {
        // Evict per-module stat entries when their document closes so _stats does
        // not grow unbounded over a long-lived window.
        this._disposables.push(
            vscode.workspace.onDidCloseTextDocument((doc) => {
                if (doc.uri.scheme === XLIDE_SCHEME) {
                    this._stats.delete(this.statKey(doc.uri));
                }
            }),
        );
    }

    // ------------------------------------------------------------------
    // Required by FileSystemProvider but not meaningful for our use case
    // ------------------------------------------------------------------

    watch(): vscode.Disposable {
        return new vscode.Disposable(() => { /* no-op */ });
    }

    stat(uri: vscode.Uri): vscode.FileStat {
        const state = this.ensureStat(uri);
        this.syncWithProjectFile(state, uri);
        return {
            type: vscode.FileType.File,
            ctime: state.ctime,
            mtime: state.mtime,
            size: state.size,
        };
    }

    readDirectory(_uri: vscode.Uri): [string, vscode.FileType][] {
        return [];
    }

    createDirectory(_uri: vscode.Uri): void {
        throw vscode.FileSystemError.NoPermissions('XLIDE: createDirectory not supported');
    }

    delete(_uri: vscode.Uri): void {
        throw vscode.FileSystemError.NoPermissions('XLIDE: delete not supported via file system');
    }

    rename(_oldUri: vscode.Uri, _newUri: vscode.Uri): void {
        throw vscode.FileSystemError.NoPermissions('XLIDE: rename not supported via file system');
    }

    // ------------------------------------------------------------------
    // Core read/write
    // ------------------------------------------------------------------

    async readFile(uri: vscode.Uri): Promise<Uint8Array> {
        const { projectPath, moduleName, face } = decodeModuleUri(uri);
        const trace = startPerformanceTrace('filesystem.readFile', moduleName);
        try {
            // A .form document is the designer's markup projection; .bas is
            // the module's code. Same project, two faces.
            const result = face === 'form'
                ? { source: (await this._bridge.call<{ markup: string }>(
                    'readFormMarkup',
                    { path: projectPath, module: moduleName },
                )).markup }
                : await this._bridge.call<{ source: string }>(
                    'readModule',
                    { path: projectPath, module: moduleName },
                );
            const bytes = Buffer.from(result.source, 'utf-8');
            this.updateSize(uri, bytes.byteLength);
            trace.end('ok', moduleName);
            return bytes;
        } catch (err) {
            trace.end('failed', moduleName);
            const message = errorMessage(err);
            if (isProjectLockedError(message)) {
                // VS Code shows its own "Unable to open" notification (with a Retry)
                // when this FileSystemError is thrown, so we do NOT also raise our
                // own warning here, which would double the popup. The thrown message
                // carries the friendly, XLIDE-prefixed guidance.
                throw vscode.FileSystemError.Unavailable(
                    `XLIDE: "${path.basename(projectPath)}" is open in ${containerAppNameForPath(projectPath)}. Close it and click Retry.`,
                );
            }
            throw err;
        }
    }

    /**
     * Saves a `.form` document: the whole edited markup goes to the engine,
     * which parses it entirely first (a parse error applies nothing), diffs
     * it against the designer by control name, and writes only what changed.
     */
    private async applyFormMarkupDocument(
        uri: vscode.Uri,
        projectPath: string,
        moduleName: string,
        markup: string,
    ): Promise<void> {
        const trace = startPerformanceTrace('filesystem.applyFormMarkup', moduleName);
        try {
            const result = await runWriteWithExcelCoordination(projectPath, () =>
                this._bridge.call<{ ok: boolean; signatureDropped: boolean; applied: string[] }>(
                    'applyFormMarkup',
                    { path: projectPath, module: moduleName, markup },
                ),
            );
            notifySignatureDropped(projectPath, result.signatureDropped);
            this.updateSize(uri, Buffer.byteLength(markup, 'utf-8'));
            recordXlideWriteAudit({
                timestamp: new Date().toISOString(),
                command: 'xlide.editorSave',
                operation: 'apply-form-markup',
                outcome: 'succeeded',
                projectPath: projectPath,
                moduleName,
                summary: result.applied.length
                    ? `Apply form markup: ${result.applied.join('; ')}`
                    : 'Apply form markup: no changes',
            });
            trace.end('ok', moduleName);
        } catch (err) {
            trace.end('failed', moduleName);
            recordXlideWriteAudit({
                timestamp: new Date().toISOString(),
                command: 'xlide.editorSave',
                operation: 'apply-form-markup',
                outcome: 'failed',
                projectPath: projectPath,
                moduleName,
                summary: `Apply form markup failed: ${errorMessage(err)}`,
            });
            // Surface the engine's own message (line-numbered for markup
            // errors) instead of a generic save failure.
            void vscode.window.showErrorMessage(`XLIDE: ${errorMessage(err)}`);
            throw vscode.FileSystemError.Unavailable(errorMessage(err));
        }
    }

    async writeFile(
        uri: vscode.Uri,
        content: Uint8Array,
        _options: { create: boolean; overwrite: boolean },
    ): Promise<void> {
        const source = Buffer.from(content).toString('utf-8');
        const { projectPath, moduleName, face } = decodeModuleUri(uri);
        if (face === 'form') {
            await this.applyFormMarkupDocument(uri, projectPath, moduleName, source);
            return;
        }
        // A rename's own edits and a developer pressing Save arrive here alike.
        // The rename registers the writes it is about to cause; anything else
        // means its before-images are stale and must not be restored over the
        // change that just happened.
        noteModuleWrite(projectPath, moduleName);
        const trace = startPerformanceTrace('filesystem.writeFile', moduleName);
        try {
            const result = await runWriteWithExcelCoordination(projectPath, () =>
                this._bridge.call<{ ok: boolean; signatureDropped: boolean }>(
                    'writeModule',
                    {
                        path: projectPath,
                        module: moduleName,
                        source,
                    },
                ),
            );
            notifySignatureDropped(projectPath, result.signatureDropped);
            // A save over a pending agent review keeps the review tracking the
            // live content - editor edits an agent makes arrive here too - so
            // Revert stays offered, still restoring the pre-agent original.
            trackModuleWriteForAgentReview(projectPath, moduleName, source);
            const summary = formatChangeSummary({
                operation: 'Save module',
                changed: [moduleName],
            });
            recordXlideWriteAudit({
                timestamp: new Date().toISOString(),
                command: 'xlide.editorSave',
                operation: 'write-module',
                outcome: 'succeeded',
                projectPath: projectPath,
                moduleName,
                summary,
            });
        } catch (err) {
            trace.end('failed', moduleName);
            const message = errorMessage(err);
            recordXlideWriteAudit({
                timestamp: new Date().toISOString(),
                command: 'xlide.editorSave',
                operation: 'write-module',
                outcome: 'failed',
                projectPath: projectPath,
                moduleName,
                summary: 'Save module: 0 changed, 1 failed',
                errorCategory: errorCategoryForSupportLog(err),
            });
            if (isProjectLockedError(message)) {
                // VS Code shows its own "Failed to save" notification (with a Retry)
                // when this FileSystemError is thrown, so we do NOT also raise our
                // own warning here, which would double the popup. The thrown message
                // carries the friendly, XLIDE-prefixed guidance.
                throw vscode.FileSystemError.Unavailable(
                    `XLIDE: "${path.basename(projectPath)}" is open in ${containerAppNameForPath(projectPath)}. Close it and save again.`,
                );
            }
            throw err;
        }
        this.markChanged(uri, Buffer.byteLength(source, 'utf-8'));
        this._emitter.fire([{ type: vscode.FileChangeType.Changed, uri }]);
        trace.end('ok', moduleName);
    }

    // Public method for out-of-band mutators (agent tools, commands)
    // to notify that a module changed so open editors reload and stats refresh
    notifyFileChanged(uri: vscode.Uri): void {
        this.markChanged(uri);
        this._emitter.fire([{ type: vscode.FileChangeType.Changed, uri }]);
    }

    dispose(): void {
        for (const d of this._disposables.splice(0)) {
            d.dispose();
        }
        this._emitter.dispose();
    }

    private ensureStat(uri: vscode.Uri): { ctime: number; mtime: number; size: number; projectKey?: string } {
        const key = this.statKey(uri);
        const existing = this._stats.get(key);
        if (existing) {
            return existing;
        }
        const real = this.projectFileMtime(uri);
        const now = real?.mtime ?? this.nextTimestamp();
        const created = { ctime: now, mtime: now, size: 0, projectKey: real?.projectKey };
        this._stats.set(key, created);
        if (real && !this._projectMtimes.has(real.projectKey)) {
            this._projectMtimes.set(real.projectKey, real.mtime);
        }
        return created;
    }

    /**
     * Real mtime of the backing project file, keyed by project identity.
     * Module mtimes are derived from it so VS Code's save-conflict detection
     * sees out-of-band changes (Excel VBE edits, module sync, agent writes).
     * Undefined for paths that cannot be statted.
     */
    private projectFileMtime(uri: vscode.Uri): { projectKey: string; mtime: number } | undefined {
        try {
            const { projectPath } = decodeModuleUri(uri);
            return {
                projectKey: projectIdentityKey(projectPath),
                mtime: Math.floor(fs.statSync(projectPath).mtimeMs),
            };
        } catch {
            return undefined;
        }
    }

    /**
     * Detect project file changes made outside this provider. When the real
     * file mtime moved past the last mtime this provider produced or observed,
     * any module may differ, so every cached module stat of that project is
     * bumped to the new file mtime.
     */
    private syncWithProjectFile(state: { mtime: number; projectKey?: string }, uri: vscode.Uri): void {
        const real = this.projectFileMtime(uri);
        if (!real) {
            return;
        }
        state.projectKey = real.projectKey;
        const baseline = this._projectMtimes.get(real.projectKey);
        if (baseline !== undefined && real.mtime !== baseline) {
            for (const entry of this._stats.values()) {
                if (entry.projectKey === real.projectKey) {
                    entry.mtime = real.mtime;
                }
            }
        }
        this._projectMtimes.set(real.projectKey, real.mtime);
    }

    private updateSize(uri: vscode.Uri, size: number): void {
        this.ensureStat(uri).size = size;
    }

    private bumpStat(uri: vscode.Uri, size?: number): void {
        const state = this.ensureStat(uri);
        state.mtime = this.nextTimestamp(state.mtime);
        if (size !== undefined) {
            state.size = size;
        }
    }

    private markChanged(uri: vscode.Uri, size?: number): void {
        const state = this.ensureStat(uri);
        const real = this.projectFileMtime(uri);
        if (real) {
            state.projectKey = real.projectKey;
            state.mtime = real.mtime;
            this._projectMtimes.set(real.projectKey, real.mtime);
        } else {
            state.mtime = this.nextTimestamp(state.mtime);
        }
        if (size !== undefined) {
            state.size = size;
        }
    }

    private nextTimestamp(after = 0): number {
        const now = Date.now();
        this._clock = Math.max(this._clock + 1, now, after + 1);
        return this._clock;
    }

    private statKey(uri: vscode.Uri): string {
        return uri.toString();
    }
}
