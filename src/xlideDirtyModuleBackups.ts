import * as crypto from 'crypto';
import * as fs from 'fs';
import * as path from 'path';
import * as vscode from 'vscode';
import { decodeModuleUri, isLocalXlideDocument } from './xlideFileSystem';
import { errorMessage } from './util/errors';

interface DirtyModuleBackup {
    uri: string;
    text: string;
    updatedAt: number;
}

interface PendingDirtyModuleBackup {
    document: vscode.TextDocument;
    timer: ReturnType<typeof setTimeout>;
    generation: number;
}

const DIRTY_BACKUP_DEBOUNCE_MS = 250;
// Backups for workbooks that are deleted/moved (or modules that are renamed)
// while dirty never get reopened, so their files are pruned once stale.
const DIRTY_BACKUP_RETENTION_MS = 30 * 24 * 60 * 60 * 1000;

function wholeDocumentRange(document: vscode.TextDocument): vscode.Range {
    const last = document.lineAt(Math.max(0, document.lineCount - 1));
    return new vscode.Range(new vscode.Position(0, 0), last.range.end);
}

function backupName(uri: vscode.Uri): string {
    return `${crypto.createHash('sha256').update(uri.toString()).digest('hex')}.json`;
}

export class XlideDirtyModuleBackups implements vscode.Disposable {
    private readonly _dir: string;
    private readonly _dirReady: Promise<void>;
    private readonly _disposables: vscode.Disposable[] = [];
    private readonly _restoring = new Set<string>();
    private readonly _announced = new Set<string>();
    private readonly _pendingWrites = new Map<string, PendingDirtyModuleBackup>();
    private readonly _writeGenerations = new Map<string, number>();
    private readonly _fileOperations = new Map<string, Promise<void>>();

    constructor(
        context: vscode.ExtensionContext,
        private readonly _out: vscode.OutputChannel,
    ) {
        this._dir = path.join(context.globalStorageUri.fsPath, 'dirty-vba-modules');
        this._dirReady = fs.promises.mkdir(this._dir, { recursive: true }).then(
            () => undefined,
            (err) => {
                const message = errorMessage(err);
                this._out.appendLine(`XLIDE: Failed to create dirty backup directory: ${message}`);
            },
        );

        this._disposables.push(
            vscode.workspace.onDidChangeTextDocument((event) => this.onDocumentChanged(event.document)),
            vscode.workspace.onDidSaveTextDocument((document) => {
                this.clearPendingWrite(document.uri);
                void this.deleteBackup(document.uri);
            }),
            vscode.workspace.onDidCloseTextDocument((document) => this.flushPendingWrite(document.uri)),
            vscode.workspace.onDidOpenTextDocument((document) => {
                void this.restoreIfAvailable(document);
            }),
        );

        for (const document of vscode.workspace.textDocuments) {
            void this.restoreIfAvailable(document);
        }

        void this.pruneStaleBackups();
    }

    dispose(): void {
        for (const disposable of this._disposables.splice(0)) {
            disposable.dispose();
        }
        for (const pending of this._pendingWrites.values()) {
            clearTimeout(pending.timer);
        }
        this._pendingWrites.clear();
    }

    private onDocumentChanged(document: vscode.TextDocument): void {
        if (!isLocalXlideDocument(document)) {
            return;
        }
        if (!document.isDirty && !this._restoring.has(document.uri.toString())) {
            this.clearPendingWrite(document.uri);
            void this.deleteBackup(document.uri);
            return;
        }
        this.scheduleWriteBackup(document);
    }

    private async restoreIfAvailable(document: vscode.TextDocument): Promise<void> {
        if (!isLocalXlideDocument(document) || document.isDirty) {
            return;
        }
        const backup = await this.readBackup(document.uri);
        if (!backup) {
            return;
        }
        const key = document.uri.toString();
        if (backup.uri !== key) {
            void this.deleteBackup(document.uri);
            return;
        }
        if (backup.text === document.getText()) {
            void this.deleteBackup(document.uri);
            return;
        }

        this._restoring.add(key);
        try {
            const edit = new vscode.WorkspaceEdit();
            edit.replace(document.uri, wholeDocumentRange(document), backup.text);
            const applied = await vscode.workspace.applyEdit(edit);
            if (!applied) {
                this._out.appendLine(`XLIDE: Could not restore unsaved backup for ${document.uri.toString()}.`);
                return;
            }
            this.announceRestore(document);
        } catch (err) {
            const message = errorMessage(err);
            this._out.appendLine(`XLIDE: Failed to restore unsaved backup for ${document.uri.toString()}: ${message}`);
        } finally {
            this._restoring.delete(key);
        }
    }

    private announceRestore(document: vscode.TextDocument): void {
        const key = document.uri.toString();
        if (this._announced.has(key)) {
            return;
        }
        this._announced.add(key);
        let label = 'module';
        try {
            const decoded = decodeModuleUri(document.uri);
            label = decoded.moduleName;
        } catch {
            /* keep generic label */
        }
        void vscode.window.showWarningMessage(
            `XLIDE restored unsaved edits for "${label}". Save to write them to the workbook, or revert the file to discard them.`,
            'Save Now',
            'Revert',
        ).then((choice) => {
            if (choice === 'Save Now') {
                void document.save();
            } else if (choice === 'Revert') {
                void vscode.window.showTextDocument(document, { preserveFocus: true }).then(() =>
                    vscode.commands.executeCommand('workbench.action.files.revert'),
                );
            }
        });
    }

    private scheduleWriteBackup(document: vscode.TextDocument): void {
        const key = document.uri.toString();
        const existing = this._pendingWrites.get(key);
        if (existing) {
            clearTimeout(existing.timer);
        }

        const generation = this.bumpWriteGeneration(key);
        // Snapshot the text inside the debounce callback so fast typing does
        // not materialize a full document copy per keystroke.
        const timer = setTimeout(() => {
            const pending = this._pendingWrites.get(key);
            if (!pending || pending.generation !== generation) {
                return;
            }
            this._pendingWrites.delete(key);
            void this.writeBackupRecord(document.uri, this.backupRecordFor(pending.document), pending.generation);
        }, DIRTY_BACKUP_DEBOUNCE_MS);
        this._pendingWrites.set(key, { document, timer, generation });
    }

    private backupRecordFor(document: vscode.TextDocument): DirtyModuleBackup {
        return {
            uri: document.uri.toString(),
            text: document.getText(),
            updatedAt: Date.now(),
        };
    }

    private flushPendingWrite(uri: vscode.Uri): void {
        const key = uri.toString();
        const pending = this._pendingWrites.get(key);
        if (!pending) {
            return;
        }
        clearTimeout(pending.timer);
        this._pendingWrites.delete(key);
        void this.writeBackupRecord(uri, this.backupRecordFor(pending.document), pending.generation);
    }

    private clearPendingWrite(uri: vscode.Uri): void {
        const key = uri.toString();
        const pending = this._pendingWrites.get(key);
        if (pending) {
            clearTimeout(pending.timer);
            this._pendingWrites.delete(key);
        }
        this.bumpWriteGeneration(key);
    }

    private async writeBackupRecord(
        uri: vscode.Uri,
        record: DirtyModuleBackup,
        generation: number,
    ): Promise<void> {
        const key = uri.toString();
        await this.enqueueFileOperation(key, async () => {
            if (this.currentWriteGeneration(key) !== generation) {
                return;
            }
            try {
                await fs.promises.writeFile(this.backupPath(uri), `${JSON.stringify(record)}\n`, 'utf8');
            } catch (err) {
                const message = errorMessage(err);
                this._out.appendLine(`XLIDE: Failed to write dirty backup for ${uri.toString()}: ${message}`);
            }
        });
    }

    private async readBackup(uri: vscode.Uri): Promise<DirtyModuleBackup | undefined> {
        try {
            const raw = await fs.promises.readFile(this.backupPath(uri), 'utf8');
            const parsed = JSON.parse(raw) as Partial<DirtyModuleBackup>;
            if (
                typeof parsed.uri !== 'string' ||
                typeof parsed.text !== 'string' ||
                typeof parsed.updatedAt !== 'number'
            ) {
                return undefined;
            }
            return parsed as DirtyModuleBackup;
        } catch {
            return undefined;
        }
    }

    private async deleteBackup(uri: vscode.Uri): Promise<void> {
        const key = uri.toString();
        this.bumpWriteGeneration(key);
        await this.enqueueFileOperation(key, async () => {
            try {
                await fs.promises.rm(this.backupPath(uri), { force: true });
            } catch {
                /* best effort cleanup */
            }
        });
    }

    private async pruneStaleBackups(): Promise<void> {
        await this._dirReady;
        let entries: string[];
        try {
            entries = await fs.promises.readdir(this._dir);
        } catch {
            return;
        }
        const cutoff = Date.now() - DIRTY_BACKUP_RETENTION_MS;
        for (const entry of entries) {
            // Skip backups owned by the restore/write flow (an open document or an
            // in-flight restore); re-checked again below after the read await.
            if (!entry.endsWith('.json') || this._isBackupInUse(entry)) {
                continue;
            }
            const fullPath = path.join(this._dir, entry);
            try {
                const raw = await fs.promises.readFile(fullPath, 'utf8');
                let updatedAt: unknown;
                try {
                    updatedAt = (JSON.parse(raw) as Partial<DirtyModuleBackup>).updatedAt;
                } catch {
                    updatedAt = undefined;
                }
                if (typeof updatedAt === 'number' && updatedAt >= cutoff) {
                    continue;
                }
                // Re-check after the await: a document may have opened and begun
                // restoring from this backup while we read it.
                if (this._isBackupInUse(entry)) {
                    continue;
                }
                await fs.promises.rm(fullPath, { force: true });
            } catch {
                /* best effort cleanup */
            }
        }
    }

    /**
     * True when a backup file is owned by an open XLIDE document or an in-flight
     * restore, so pruneStaleBackups must not delete it out from under that flow.
     */
    private _isBackupInUse(backupFileName: string): boolean {
        for (const document of vscode.workspace.textDocuments) {
            if (isLocalXlideDocument(document) && backupName(document.uri) === backupFileName) {
                return true;
            }
        }
        for (const uriString of this._restoring) {
            try {
                if (backupName(vscode.Uri.parse(uriString)) === backupFileName) {
                    return true;
                }
            } catch {
                /* unparsable restoring key - ignore */
            }
        }
        return false;
    }

    private bumpWriteGeneration(key: string): number {
        const next = this.currentWriteGeneration(key) + 1;
        this._writeGenerations.set(key, next);
        return next;
    }

    private currentWriteGeneration(key: string): number {
        return this._writeGenerations.get(key) ?? 0;
    }

    private enqueueFileOperation(key: string, operation: () => Promise<void>): Promise<void> {
        const previous = this._fileOperations.get(key) ?? this._dirReady;
        const next = previous
            .catch(() => {
                /* each operation handles its own reporting */
            })
            .then(operation)
            .finally(() => {
                if (this._fileOperations.get(key) === next) {
                    this._fileOperations.delete(key);
                }
            });
        this._fileOperations.set(key, next);
        return next;
    }

    private backupPath(uri: vscode.Uri): string {
        return path.join(this._dir, backupName(uri));
    }
}

export function registerXlideDirtyModuleBackups(
    context: vscode.ExtensionContext,
    out: vscode.OutputChannel,
): vscode.Disposable {
    return new XlideDirtyModuleBackups(context, out);
}
