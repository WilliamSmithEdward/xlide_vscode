import * as vscode from 'vscode';
import {
    errorCategoryForSupportLog,
    recordXlideCommand,
} from './xlideCommandLog';
import { startPerformanceTrace } from './performanceTrace';
import { errorMessage } from './util/errors';

export interface XlideCommandErrorOptions<T extends unknown[]> {
    /** Shown to the user as `XLIDE: ${errorPrefix}: ${message}`. */
    errorPrefix: string;
    /** Output-channel tag; failures log as `[logTag] Error: ${message}`. */
    logTag: string;
    log: (message: string) => void;
    /** Extra failure handling (e.g. write-audit) before the message is shown. */
    onError?: (err: unknown, ...args: T) => void | Promise<void>;
}

export function registerXlideCommand<T extends unknown[]>(
    command: string,
    callback: (...args: T) => unknown,
    errorOptions?: XlideCommandErrorOptions<T>,
): vscode.Disposable {
    return vscode.commands.registerCommand(command, async (...args: T) => {
        const start = Date.now();
        const trace = startPerformanceTrace('command', command);
        recordXlideCommand({
            timestamp: new Date(start).toISOString(),
            command,
            outcome: 'started',
        });
        try {
            const result = await callback(...args);
            trace.end('ok', command);
            recordXlideCommand({
                timestamp: new Date().toISOString(),
                command,
                outcome: 'succeeded',
                durationMs: Date.now() - start,
            });
            return result;
        } catch (err) {
            const canceled = err instanceof vscode.CancellationError;
            trace.end(canceled ? 'canceled' : 'failed', command);
            recordXlideCommand({
                timestamp: new Date().toISOString(),
                command,
                outcome: 'failed',
                durationMs: Date.now() - start,
                errorCategory: errorCategoryForSupportLog(err),
            });
            if (!errorOptions) {
                throw err;
            }
            if (canceled) {
                return undefined;
            }
            const message = errorMessage(err);
            errorOptions.log(`[${errorOptions.logTag}] Error: ${message}`);
            await errorOptions.onError?.(err, ...args);
            vscode.window.showErrorMessage(`XLIDE: ${errorOptions.errorPrefix}: ${message}`);
            return undefined;
        }
    });
}
