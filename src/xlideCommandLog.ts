import { errorMessage } from './util/errors';
export interface XlideCommandLogEntry {
    timestamp: string;
    command: string;
    outcome: 'started' | 'succeeded' | 'failed';
    durationMs?: number;
    errorCategory?: string;
}

const MAX_COMMAND_LOG_ENTRIES = 100;
const entries: XlideCommandLogEntry[] = [];

export function recordXlideCommand(entry: XlideCommandLogEntry): void {
    entries.push(entry);
    if (entries.length > MAX_COMMAND_LOG_ENTRIES) {
        entries.splice(0, entries.length - MAX_COMMAND_LOG_ENTRIES);
    }
}

export function recentXlideCommands(limit = 25): XlideCommandLogEntry[] {
    return entries.slice(Math.max(0, entries.length - limit));
}

export function clearXlideCommandLog(): void {
    entries.splice(0);
}

/**
 * Error signatures indicating Excel holds the workbook open (Windows file
 * sharing violation).  Shared by support-log categorization here and the
 * user-facing project-locked warning in xlideFileSystem.
 */
export const WORKBOOK_LOCKED_ERROR_RE =
    /WinError\s*3[23]\b|WinError\s*5\b|being used by another process|sharing violation|access is denied|permission denied|permissionerror|\bEACCES\b|\bEBUSY\b/i;

export function errorCategoryForSupportLog(error: unknown): string {
    const message = errorMessage(error);
    if (WORKBOOK_LOCKED_ERROR_RE.test(message)) {
        return 'project-locked';
    }
    if (/\bENOENT\b|no such file|cannot find|not found/i.test(message)) {
        return 'project-missing';
    }
    if (/cancel|canceled|cancelled/i.test(message)) {
        return 'cancelled';
    }
    return 'unknown';
}
