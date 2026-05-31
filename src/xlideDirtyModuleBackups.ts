import * as crypto from 'crypto';
import * as fs from 'fs';
import * as path from 'path';
import * as vscode from 'vscode';
import { decodeModuleUri, XLIDE_LIVESHARE_AUTHORITY, XLIDE_SCHEME } from './xlideFileSystem';

interface DirtyModuleBackup {
    uri: string;
    text: string;
    updatedAt: number;
}

function isLocalXlideModule(document: vscode.TextDocument): boolean {
    return document.uri.scheme === XLIDE_SCHEME && document.uri.authority !== XLIDE_LIVESHARE_AUTHORITY;
}

function wholeDocumentRange(document: vscode.TextDocument): vscode.Range {
    const last = document.lineAt(Math.max(0, document.lineCount - 1));
    return new vscode.Range(new vscode.Position(0, 0), last.range.end);
}

function backupName(uri: vscode.Uri): string {
    return `${crypto.createHash('sha256').update(uri.toString()).digest('hex')}.json`;
}

export class XlideDirtyModuleBackups implements vscode.Disposable {
    private readonly _dir: string;
    private readonly _disposables: vscode.Disposable[] = [];
    private readonly _restoring = new Set<string>();
    private readonly _announced = new Set<string>();

    constructor(
        context: vscode.ExtensionContext,
        private readonly _out: vscode.OutputChannel,
    ) {
        this._dir = path.join(context.globalStorageUri.fsPath, 'dirty-vba-modules');
        fs.mkdirSync(this._dir, { recursive: true });

        this._disposables.push(
            vscode.workspace.onDidChangeTextDocument((event) => this.onDocumentChanged(event.document)),
            vscode.workspace.onDidSaveTextDocument((document) => this.deleteBackup(document.uri)),
            vscode.workspace.onDidOpenTextDocument((document) => {
                void this.restoreIfAvailable(document);
            }),
        );

        for (const document of vscode.workspace.textDocuments) {
            void this.restoreIfAvailable(document);
        }
    }

    dispose(): void {
        for (const disposable of this._disposables.splice(0)) {
            disposable.dispose();
        }
    }

    private onDocumentChanged(document: vscode.TextDocument): void {
        if (!isLocalXlideModule(document)) {
            return;
        }
        if (!document.isDirty && !this._restoring.has(document.uri.toString())) {
            this.deleteBackup(document.uri);
            return;
        }
        this.writeBackup(document);
    }

    private async restoreIfAvailable(document: vscode.TextDocument): Promise<void> {
        if (!isLocalXlideModule(document) || document.isDirty) {
            return;
        }
        const backup = this.readBackup(document.uri);
        if (!backup) {
            return;
        }
        const key = document.uri.toString();
        if (backup.uri !== key) {
            this.deleteBackup(document.uri);
            return;
        }
        if (backup.text === document.getText()) {
            this.deleteBackup(document.uri);
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
            const message = err instanceof Error ? err.message : String(err);
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

    private writeBackup(document: vscode.TextDocument): void {
        try {
            const record: DirtyModuleBackup = {
                uri: document.uri.toString(),
                text: document.getText(),
                updatedAt: Date.now(),
            };
            fs.writeFileSync(this.backupPath(document.uri), `${JSON.stringify(record)}\n`, 'utf8');
        } catch (err) {
            const message = err instanceof Error ? err.message : String(err);
            this._out.appendLine(`XLIDE: Failed to write dirty backup for ${document.uri.toString()}: ${message}`);
        }
    }

    private readBackup(uri: vscode.Uri): DirtyModuleBackup | undefined {
        try {
            const raw = fs.readFileSync(this.backupPath(uri), 'utf8');
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

    private deleteBackup(uri: vscode.Uri): void {
        try {
            fs.rmSync(this.backupPath(uri), { force: true });
        } catch {
            /* best effort cleanup */
        }
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
