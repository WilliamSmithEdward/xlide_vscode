// Existence check for the web extension host, asked through the editor.
//
// A web workspace's files live behind a FileSystemProvider - github.dev's
// vscode-vfs://, for instance - that only VS Code can read, so workspace.fs
// is the only thing that can answer. It works for any scheme, local included.

import * as vscode from 'vscode';
import { workspaceUriFor } from './workspaceUris';

export async function fileExists(filePath: string): Promise<boolean> {
    try {
        await vscode.workspace.fs.stat(workspaceUriFor(filePath));
        return true;
    } catch {
        return false;
    }
}
