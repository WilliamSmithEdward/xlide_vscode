import * as vscode from 'vscode';
import {
    errorCategoryForSupportLog,
    recordXlideCommand,
} from './xlideCommandLog';

export function registerXlideCommand<T extends unknown[]>(
    command: string,
    callback: (...args: T) => unknown,
): vscode.Disposable {
    return vscode.commands.registerCommand(command, async (...args: T) => {
        const start = Date.now();
        recordXlideCommand({
            timestamp: new Date(start).toISOString(),
            command,
            outcome: 'started',
        });
        try {
            const result = await callback(...args);
            recordXlideCommand({
                timestamp: new Date().toISOString(),
                command,
                outcome: 'succeeded',
                durationMs: Date.now() - start,
            });
            return result;
        } catch (err) {
            recordXlideCommand({
                timestamp: new Date().toISOString(),
                command,
                outcome: 'failed',
                durationMs: Date.now() - start,
                errorCategory: errorCategoryForSupportLog(err),
            });
            throw err;
        }
    });
}
