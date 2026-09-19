import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';
import { ProjectEngine } from './projectEngine';
import { errorCategoryForSupportLog, PROJECT_LOCKED_ERROR_RE } from './xlideCommandLog';
import { formatChangeSummary, recordXlideWriteAudit } from './xlideWriteAudit';
import { startPerformanceTrace } from './performanceTrace';
import { errorMessage } from './util/errors';
import { runWriteWithHostCoordination } from './officeWriteCoordinator';
import { noteModuleWrite } from './vbaRenameHistory';
// Function-level cycle with xlideAgentDiff (it imports URI/identity helpers
// from this module); neither side touches the other at module-eval time.
import { trackModuleWriteForAgentReview } from './xlideAgentDiff';
import { containerAppNameForPath, MACRO_CONTAINER_EXTENSION_PATTERN } from './macroContainerUi';
import { projectIdentityKey } from './projectIdentity';
import { checkProjectFile, onDidChangeProjectFile, watchProjectFile } from './projectFileChanges';
import { moduleContentToken } from './moduleContentToken';

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
 * Heuristic: does this error string look like the file's application (Excel,
 * Word, PowerPoint or Access) holding it open?
 */
export function isProjectLockedError(message: string): boolean {
    return PROJECT_LOCKED_ERROR_RE.test(message);
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

/** How long a closed document stays closed before its entry goes. */
const FORGET_CLOSED_AFTER_MS = 1000;

/** What the provider knows of one module document it has served. */
interface ModuleStat {
    ctime: number;
    mtime: number;
    size: number;
    /** What the document was last given, by a read or by its own save. */
    given?: { token: string; size: number };
    /**
     * When the entry was made or last given content. VS Code lists a document
     * as open only after reading it, so a sweep leaves a recent entry be.
     */
    touched: number;
}

/**
 * The size to report for a module whose content moved away from what its
 * document was given. VS Code stops a save at "File Modified Since" only when
 * the size differs as well as the mtime, so this never equals the size the
 * document was given: the new content's own size when that differs, else one
 * byte more. A reload replaces it with the real size.
 */
function sizeAfterChange(state: ModuleStat, newSize?: number): number {
    const givenSize = state.given?.size ?? state.size;
    return newSize !== undefined && newSize !== givenSize ? newSize : givenSize + 1;
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
    private readonly _stats = new Map<string, ModuleStat>();
    /** A watch on each project file while a module of it is open. */
    private readonly _projectWatches = new Map<string, vscode.Disposable>();
    /** Open modules being compared with a project file that changed outside XLIDE. */
    private readonly _reconciles = new Map<string, Promise<void>>();
    private _forgetTimer: ReturnType<typeof setTimeout> | undefined;
    private readonly _disposables: vscode.Disposable[] = [];

    constructor(private readonly _bridge: ProjectEngine) {
        this._disposables.push(
            vscode.workspace.onDidOpenTextDocument((doc) => {
                if (doc.uri.scheme === XLIDE_SCHEME) {
                    this.watchOpenProjects();
                }
            }),
            // Evict per-module stat entries when their document closes so _stats does
            // not grow unbounded over a long-lived window. Not at once: setting a
            // document's language closes it and opens it again, and its entry
            // has to survive that.
            vscode.workspace.onDidCloseTextDocument((doc) => {
                if (doc.uri.scheme === XLIDE_SCHEME) {
                    this.scheduleForgetClosed();
                }
            }),
            onDidChangeProjectFile((projectPath) => this.reconcile(projectPath)),
        );
        this.watchOpenProjects();
    }

    // ------------------------------------------------------------------
    // Required by FileSystemProvider but not meaningful for our use case
    // ------------------------------------------------------------------

    watch(): vscode.Disposable {
        return new vscode.Disposable(() => { /* no-op */ });
    }

    async stat(uri: vscode.Uri): Promise<vscode.FileStat> {
        const project = this.projectFileMtime(uri);
        if (project) {
            // Settle a change made outside XLIDE before answering, so a save
            // cannot pass over it: the modules it reached answer with a newer
            // mtime, and VS Code reports the conflict.
            checkProjectFile(project.projectPath);
            await this._reconciles.get(project.projectKey);
        }
        const state = this.ensureStat(uri);
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
            const source = await this.readContent(projectPath, moduleName, face);
            const bytes = Buffer.from(source, 'utf-8');
            this.recordContent(uri, source, bytes.byteLength);
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
            // A module the project no longer has: VS Code keeps a document a
            // while after its tab closes, and reads it again when the module is
            // deleted or renamed. It takes FileNotFound as the file being gone,
            // where it logs any other error as the provider failing.
            if (/^Module not found: /.test(message)) {
                throw vscode.FileSystemError.FileNotFound(uri);
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
            const result = await runWriteWithHostCoordination(projectPath, () =>
                this._bridge.call<{ ok: boolean; signatureDropped: boolean; applied: string[] }>(
                    'applyFormMarkup',
                    { path: projectPath, module: moduleName, markup },
                ),
            );
            notifySignatureDropped(projectPath, result.signatureDropped);
            await this.recordSave(uri, projectPath, moduleName, 'form', markup);
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
            const result = await runWriteWithHostCoordination(projectPath, () =>
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
        await this.recordSave(uri, projectPath, moduleName, 'code', source);
        this._emitter.fire([{ type: vscode.FileChangeType.Changed, uri }]);
        trace.end('ok', moduleName);
    }

    // Public method for out-of-band mutators (agent tools, commands)
    // to notify that a module changed so open editors reload and stats refresh.
    // The stat moves past what the document holds, so a document with unsaved
    // edits cannot save over the change without VS Code's conflict prompt.
    notifyFileChanged(uri: vscode.Uri): void {
        const state = this.ensureStat(uri);
        this.advanceMtime(state, uri);
        state.size = sizeAfterChange(state);
        this._emitter.fire([{ type: vscode.FileChangeType.Changed, uri }]);
    }

    dispose(): void {
        for (const d of this._disposables.splice(0)) {
            d.dispose();
        }
        if (this._forgetTimer) {
            clearTimeout(this._forgetTimer);
        }
        for (const watch of this._projectWatches.values()) {
            watch.dispose();
        }
        this._projectWatches.clear();
        this._emitter.dispose();
    }

    private async readContent(projectPath: string, moduleName: string, face: 'code' | 'form' | undefined): Promise<string> {
        // A .form document is the designer's markup projection; .bas is
        // the module's code. Same project, two faces.
        if (face === 'form') {
            return (await this._bridge.call<{ markup: string }>(
                'readFormMarkup',
                { path: projectPath, module: moduleName },
            )).markup;
        }
        return (await this._bridge.call<{ source: string }>(
            'readModule',
            { path: projectPath, module: moduleName },
        )).source;
    }

    private ensureStat(uri: vscode.Uri): ModuleStat {
        const key = this.statKey(uri);
        const existing = this._stats.get(key);
        if (existing) {
            return existing;
        }
        const real = this.projectFileMtime(uri);
        const now = real?.mtime ?? this.nextTimestamp();
        const created: ModuleStat = { ctime: now, mtime: now, size: 0, touched: Date.now() };
        this._stats.set(key, created);
        return created;
    }

    /** The module documents VS Code has open, with the project each belongs to. */
    private openModuleDocuments(): Array<{ document: vscode.TextDocument; projectPath: string }> {
        const out: Array<{ document: vscode.TextDocument; projectPath: string }> = [];
        for (const document of vscode.workspace.textDocuments) {
            if (document.uri.scheme !== XLIDE_SCHEME || document.isClosed) {
                continue;
            }
            try {
                out.push({ document, projectPath: decodeModuleUri(document.uri).projectPath });
            } catch {
                // Not a module address; nothing to watch.
            }
        }
        return out;
    }

    /** Watches the file of every project with a module open, and only those. */
    private watchOpenProjects(): void {
        const open = new Map<string, string>();
        for (const { projectPath } of this.openModuleDocuments()) {
            open.set(projectIdentityKey(projectPath), projectPath);
        }
        for (const [projectKey, projectPath] of open) {
            if (!this._projectWatches.has(projectKey)) {
                this._projectWatches.set(projectKey, watchProjectFile(projectPath));
            }
        }
        for (const [projectKey, watch] of this._projectWatches) {
            if (!open.has(projectKey)) {
                watch.dispose();
                this._projectWatches.delete(projectKey);
            }
        }
    }

    /**
     * Once closes have settled, drops the entries of documents that stayed
     * closed. An entry made or read within the wait may belong to a document
     * VS Code is opening and has not listed yet; it waits for the next sweep.
     */
    private scheduleForgetClosed(): void {
        if (this._forgetTimer) {
            clearTimeout(this._forgetTimer);
        }
        this._forgetTimer = setTimeout(() => {
            this._forgetTimer = undefined;
            const open = new Set(this.openModuleDocuments().map(({ document }) => this.statKey(document.uri)));
            let waiting = false;
            for (const [key, entry] of [...this._stats]) {
                if (open.has(key)) {
                    continue;
                }
                if (Date.now() - entry.touched < FORGET_CLOSED_AFTER_MS) {
                    waiting = true;
                } else {
                    this._stats.delete(key);
                }
            }
            this.watchOpenProjects();
            if (waiting) {
                this.scheduleForgetClosed();
            }
        }, FORGET_CLOSED_AFTER_MS);
    }

    /**
     * Real mtime of the backing project file, keyed by project identity.
     * Module mtimes are derived from it so VS Code's save-conflict detection
     * sees changes made outside XLIDE (Excel VBE edits, git, another window).
     * Undefined for paths that cannot be statted.
     */
    private projectFileMtime(uri: vscode.Uri): { projectPath: string; projectKey: string; mtime: number } | undefined {
        try {
            const { projectPath } = decodeModuleUri(uri);
            return {
                projectPath,
                projectKey: projectIdentityKey(projectPath),
                mtime: Math.floor(fs.statSync(projectPath).mtimeMs),
            };
        } catch {
            return undefined;
        }
    }

    /**
     * The project file changed and XLIDE did not write it. Each open module is
     * read again and compared with its document: one the change reached moves
     * its mtime forward and is reported changed, so VS Code reloads it if it
     * has no unsaved edits and reports a conflict if it has. One the change did
     * not reach keeps its mtime, so its unsaved edits still save cleanly.
     * Queued per project, so two changes never compare at once.
     */
    private reconcile(projectPath: string): void {
        const projectKey = projectIdentityKey(projectPath);
        // Taken now: a module opened later is read at the new content anyway.
        const documents = this.openModuleDocuments()
            .filter((open) => projectIdentityKey(open.projectPath) === projectKey)
            .map((open) => open.document);
        const previous = this._reconciles.get(projectKey) ?? Promise.resolve();
        const run = previous
            .then(() => this.reconcileDocuments(projectPath, documents))
            .catch(() => undefined);
        this._reconciles.set(projectKey, run);
        void run.then(() => {
            if (this._reconciles.get(projectKey) === run) {
                this._reconciles.delete(projectKey);
            }
        });
    }

    private async reconcileDocuments(projectPath: string, documents: readonly vscode.TextDocument[]): Promise<void> {
        let fileMtime = 0;
        try {
            fileMtime = Math.floor(fs.statSync(projectPath).mtimeMs);
        } catch {
            // Gone for the moment; the mtimes still move forward below.
        }
        const changed: vscode.FileChangeEvent[] = [];
        for (const document of documents) {
            const entry = this.ensureStat(document.uri);
            let current: string | undefined;
            try {
                const { moduleName, face } = decodeModuleUri(document.uri);
                current = await this.readContent(projectPath, moduleName, face);
            } catch {
                // Not readable - the module was removed, or the file is
                // mid-save - so it cannot be shown unchanged.
            }
            // What the document last loaded, or what its own save left in the
            // module; a document with no unsaved edits that the provider has no
            // record of is known by its text.
            const given = entry.given?.token ?? (document.isDirty ? undefined : moduleContentToken(document.getText()));
            if (current !== undefined && given !== undefined && moduleContentToken(current) === given) {
                continue;
            }
            entry.mtime = Math.max(fileMtime, entry.mtime + 1);
            entry.size = sizeAfterChange(entry, current === undefined ? undefined : Buffer.byteLength(current, 'utf-8'));
            changed.push({ type: vscode.FileChangeType.Changed, uri: document.uri });
        }
        if (changed.length > 0) {
            this._emitter.fire(changed);
        }
    }

    /** The text a document was just given, and its size. */
    private recordContent(uri: vscode.Uri, source: string, size: number): void {
        const state = this.ensureStat(uri);
        state.size = size;
        state.given = { token: moduleContentToken(source), size };
        state.touched = Date.now();
    }

    /**
     * The document's own save. What it was given is what the module holds
     * now, read back: the engine does not always store the text as written -
     * it drops blank lines above a module's code, and a designer's markup
     * comes back in the designer's own layout - and a record of the text as
     * written made the next change anywhere in the project look like a change
     * to this module, stopping its next save at a conflict nobody caused.
     */
    private async recordSave(
        uri: vscode.Uri,
        projectPath: string,
        moduleName: string,
        face: 'code' | 'form',
        written: string,
    ): Promise<void> {
        let stored = written;
        try {
            const readBack: unknown = await this.readContent(projectPath, moduleName, face);
            if (typeof readBack === 'string') {
                stored = readBack;
            }
        } catch {
            // The write succeeded; what was written is the best record left.
        }
        const state = this.ensureStat(uri);
        this.advanceMtime(state, uri);
        this.recordContent(uri, stored, Buffer.byteLength(stored, 'utf-8'));
    }

    private advanceMtime(state: ModuleStat, uri: vscode.Uri): void {
        const real = this.projectFileMtime(uri);
        state.mtime = real ? Math.max(real.mtime, state.mtime + 1) : this.nextTimestamp(state.mtime);
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
