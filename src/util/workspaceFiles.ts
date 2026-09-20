// Reading and writing real files in the workspace: settings, exported
// modules, backups of unsaved edits.
//
// This is not the container engine's file access - that is hostPlatform.ts
// and is synchronous by necessity. Everything here is already async, so the
// browser build can go straight to vscode.workspace.fs, which is the only
// thing that can see a virtual workspace's files. The desktop build keeps
// node:fs, so nothing about its behaviour changes.
//
// Missing files are expressed as outcomes rather than as error codes: a
// caller asking "is this settings file there" should not have to know the
// difference between ENOENT and FileSystemError.FileNotFound.

import { workspaceFiles as leaf } from './workspaceFilesNode';

export interface DirectoryEntry {
    name: string;
    isDirectory: boolean;
}

export interface WorkspaceFiles {
    /** Undefined when the file does not exist; other failures throw. */
    readTextIfPresent(filePath: string): Promise<string | undefined>;

    readBytes(filePath: string): Promise<Buffer>;

    writeText(filePath: string, text: string): Promise<void>;

    writeBytes(filePath: string, data: Buffer): Promise<void>;

    /** False when the file was already there, in which case nothing is written. */
    createTextIfAbsent(filePath: string, text: string): Promise<boolean>;

    /** Creates the directory and any missing parents. */
    makeDirectory(dirPath: string): Promise<void>;

    listDirectory(dirPath: string): Promise<DirectoryEntry[]>;

    /** Deleting something that is not there is not an error. */
    remove(filePath: string): Promise<void>;

    /**
     * Best effort, for deactivation. dispose() cannot await, so the desktop
     * writes synchronously and the browser - which has no synchronous write -
     * starts the write and lets it finish if the host lives long enough.
     */
    writeTextDuringShutdown(filePath: string, text: string): void;

    /** Best effort, paired with writeTextDuringShutdown. */
    makeDirectoryDuringShutdown(dirPath: string): void;
}

export const workspaceFiles: WorkspaceFiles = leaf;
