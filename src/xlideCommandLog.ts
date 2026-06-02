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

export function errorCategoryForSupportLog(error: unknown): string {
    const message = error instanceof Error ? error.message : String(error);
    if (/WinError\s*32|being used by another process|sharing violation|Permission denied|PermissionError/i.test(message)) {
        return 'workbook-locked';
    }
    if (/python|backend|spawn|ENOENT/i.test(message)) {
        return 'python-backend';
    }
    if (/cancel|canceled|cancelled/i.test(message)) {
        return 'cancelled';
    }
    return 'unknown';
}
