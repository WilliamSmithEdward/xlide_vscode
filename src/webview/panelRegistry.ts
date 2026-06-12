import * as vscode from 'vscode';
import { workbookIdentityKey } from '../xlideFileSystem';
import { errorMessage } from '../util/errors';

export interface WebviewPanelRegistry<TEntry> {
    /** Returns the open entry for the workbook, if any. */
    get(filePath: string): TEntry | undefined;
    /** Stores the entry; the panel's dispose handler must call delete(filePath). */
    set(filePath: string, entry: TEntry): void;
    delete(filePath: string): void;
}

/** Singleton-panel registry keyed by canonical workbook identity. */
export function createWebviewPanelRegistry<TEntry>(): WebviewPanelRegistry<TEntry> {
    const open = new Map<string, TEntry>();
    return {
        get: (filePath) => open.get(workbookIdentityKey(filePath)),
        set: (filePath, entry) => {
            open.set(workbookIdentityKey(filePath), entry);
        },
        delete: (filePath) => {
            open.delete(workbookIdentityKey(filePath));
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
