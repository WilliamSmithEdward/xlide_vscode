// Workspace file access in the browser, through vscode.workspace.fs.
//
// A web workspace's files sit behind a FileSystemProvider - github.dev's
// vscode-vfs://, for one - so the editor is the only thing that can read
// them. workspace.fs is async throughout, which suits every caller here
// except the shutdown pair; see below.

import * as vscode from 'vscode';
import type { DirectoryEntry, WorkspaceFiles } from './workspaceFiles';
import { workspaceUriFor } from './workspaceUris';

const encoder = new TextEncoder();
const decoder = new TextDecoder('utf-8');

function uriOf(filePath: string): vscode.Uri {
    return workspaceUriFor(filePath);
}

function isMissing(err: unknown): boolean {
    return err instanceof vscode.FileSystemError && err.code === 'FileNotFound';
}

export const workspaceFiles: WorkspaceFiles = {
    async readTextIfPresent(filePath: string): Promise<string | undefined> {
        try {
            return decoder.decode(await vscode.workspace.fs.readFile(uriOf(filePath)));
        } catch (err) {
            if (isMissing(err)) {
                return undefined;
            }
            throw err;
        }
    },

    async readBytes(filePath: string): Promise<Buffer> {
        return Buffer.from(await vscode.workspace.fs.readFile(uriOf(filePath)));
    },

    async writeText(filePath: string, text: string): Promise<void> {
        await vscode.workspace.fs.writeFile(uriOf(filePath), encoder.encode(text));
    },

    async writeBytes(filePath: string, data: Buffer): Promise<void> {
        await vscode.workspace.fs.writeFile(uriOf(filePath), data);
    },

    async createTextIfAbsent(filePath: string, text: string): Promise<boolean> {
        // workspace.fs has no exclusive-create flag, so this is a check and
        // then a write rather than one atomic step. The callers use it to put
        // a default settings file in place, where the loser of a race writes
        // the same bytes the winner did.
        try {
            await vscode.workspace.fs.stat(uriOf(filePath));
            return false;
        } catch (err) {
            if (!isMissing(err)) {
                throw err;
            }
        }
        await vscode.workspace.fs.writeFile(uriOf(filePath), encoder.encode(text));
        return true;
    },

    async makeDirectory(dirPath: string): Promise<void> {
        // createDirectory creates missing parents, like mkdir -p.
        await vscode.workspace.fs.createDirectory(uriOf(dirPath));
    },

    async listDirectory(dirPath: string): Promise<DirectoryEntry[]> {
        const entries = await vscode.workspace.fs.readDirectory(uriOf(dirPath));
        return entries.map(([name, type]) => ({
            name,
            isDirectory: (type & vscode.FileType.Directory) !== 0,
        }));
    },

    async remove(filePath: string): Promise<void> {
        try {
            await vscode.workspace.fs.delete(uriOf(filePath), { recursive: true });
        } catch (err) {
            if (!isMissing(err)) {
                throw err;
            }
        }
    },

    writeTextDuringShutdown(filePath: string, text: string): void {
        // There is no synchronous write in a browser. Starting it and not
        // waiting keeps the same best-effort contract the desktop has: the
        // snapshot lands if the host outlives the call, and an abrupt close
        // loses it either way.
        void vscode.workspace.fs.writeFile(uriOf(filePath), encoder.encode(text)).then(
            undefined,
            () => { /* best effort during shutdown */ },
        );
    },

    makeDirectoryDuringShutdown(dirPath: string): void {
        void vscode.workspace.fs.createDirectory(uriOf(dirPath)).then(
            undefined,
            () => { /* best effort during shutdown */ },
        );
    },
};
