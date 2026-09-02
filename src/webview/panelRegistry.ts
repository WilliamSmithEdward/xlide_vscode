import * as vscode from 'vscode';
import { projectIdentityKey } from '../xlideFileSystem';
import { errorMessage } from '../util/errors';

export interface WebviewPanelRegistry<TEntry> {
    /** Returns the open entry for the project, if any. */
    get(filePath: string): TEntry | undefined;
    /** Stores the entry; the panel's dispose handler must call delete(filePath). */
    set(filePath: string, entry: TEntry): void;
    delete(filePath: string): void;
}

/** Singleton-panel registry keyed by canonical project identity. */
export function createWebviewPanelRegistry<TEntry>(): WebviewPanelRegistry<TEntry> {
    const open = new Map<string, TEntry>();
    return {
        get: (filePath) => open.get(projectIdentityKey(filePath)),
        set: (filePath, entry) => {
            open.set(projectIdentityKey(filePath), entry);
        },
        delete: (filePath) => {
            open.delete(projectIdentityKey(filePath));
        },
    };
}

/**
 * Subscribes the handler to webview messages and surfaces handler failures
 * to the webview as a { type: 'error', error } message.
 */
export function bridgeWebviewMessages<TMessage>(
    webview: vscode.Webview,
    handler: (message: TMessage) => Promise<void>,
    onError?: (error: unknown) => Promise<void> | void,
): vscode.Disposable {
    return webview.onDidReceiveMessage(async (message: TMessage) => {
        try {
            await handler(message);
        } catch (err) {
            await webview.postMessage({ type: 'error', error: errorMessage(err) });
            await onError?.(err);
        }
    });
}
