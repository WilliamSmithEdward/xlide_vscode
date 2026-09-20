// Workspace file access on a desktop: node:fs, exactly as before the browser
// build existed. The browser never reaches this module - webBuild.js aliases
// it to workspaceFilesWeb.ts.

import * as fs from 'fs';
import type { DirectoryEntry, WorkspaceFiles } from './workspaceFiles';

function isMissing(err: unknown): boolean {
    return (err as NodeJS.ErrnoException)?.code === 'ENOENT';
}

export const workspaceFiles: WorkspaceFiles = {
    async readTextIfPresent(filePath: string): Promise<string | undefined> {
        try {
            return await fs.promises.readFile(filePath, 'utf8');
        } catch (err) {
            if (isMissing(err)) {
                return undefined;
            }
            throw err;
        }
    },

    async readBytes(filePath: string): Promise<Buffer> {
        return fs.promises.readFile(filePath);
    },

    async writeText(filePath: string, text: string): Promise<void> {
        await fs.promises.writeFile(filePath, text, 'utf8');
    },

    async writeBytes(filePath: string, data: Buffer): Promise<void> {
        await fs.promises.writeFile(filePath, data);
    },

    async createTextIfAbsent(filePath: string, text: string): Promise<boolean> {
        try {
            // 'wx' fails rather than truncating when the file is already
            // there, so two windows racing to create it cannot lose one's
            // contents.
            await fs.promises.writeFile(filePath, text, { encoding: 'utf8', flag: 'wx' });
            return true;
        } catch (err) {
            if ((err as NodeJS.ErrnoException)?.code === 'EEXIST') {
                return false;
            }
            throw err;
        }
    },

    async makeDirectory(dirPath: string): Promise<void> {
        await fs.promises.mkdir(dirPath, { recursive: true });
    },

    async listDirectory(dirPath: string): Promise<DirectoryEntry[]> {
        const entries = await fs.promises.readdir(dirPath, { withFileTypes: true });
        return entries.map((entry) => ({ name: entry.name, isDirectory: entry.isDirectory() }));
    },

    async remove(filePath: string): Promise<void> {
        await fs.promises.rm(filePath, { force: true });
    },

    writeTextDuringShutdown(filePath: string, text: string): void {
        fs.writeFileSync(filePath, text, 'utf8');
    },

    makeDirectoryDuringShutdown(dirPath: string): void {
        fs.mkdirSync(dirPath, { recursive: true });
    },
};
