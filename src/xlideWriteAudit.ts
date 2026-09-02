import { errorCategoryForSupportLog } from './xlideCommandLog';

export type XlideWriteAuditOutcome = 'succeeded' | 'failed' | 'skipped';

export interface XlideWriteAuditEntry {
    timestamp: string;
    command: string;
    operation: string;
    outcome: XlideWriteAuditOutcome;
    projectPath?: string;
    moduleName?: string;
    sourcePath?: string;
    targetPath?: string;
    summary: string;
    errorCategory?: string;
}

const MAX_WRITE_AUDIT_ENTRIES = 100;
const writeAudit: XlideWriteAuditEntry[] = [];

export function recordXlideWriteAudit(entry: XlideWriteAuditEntry): void {
    writeAudit.push({ ...entry });
    if (writeAudit.length > MAX_WRITE_AUDIT_ENTRIES) {
        writeAudit.splice(0, writeAudit.length - MAX_WRITE_AUDIT_ENTRIES);
    }
}

export function recentXlideWriteAudits(limit = 25): XlideWriteAuditEntry[] {
    if (limit <= 0) { return []; }
    return writeAudit.slice(-limit);
}

export function clearXlideWriteAudit(): void {
    writeAudit.length = 0;
}

export type XlideWriteAuditEvent = Omit<XlideWriteAuditEntry, 'timestamp' | 'errorCategory'> & {
    error?: unknown;
};

/** Records an audit entry stamped now, deriving errorCategory from `error`. */
export function recordXlideWriteAuditEvent(event: XlideWriteAuditEvent): void {
    const { error, ...entry } = event;
    recordXlideWriteAudit({
        timestamp: new Date().toISOString(),
        ...entry,
        errorCategory: error ? errorCategoryForSupportLog(error) : undefined,
    });
}

export interface XlideWriteAuditTarget {
    command: string;
    operation: string;
    projectPath?: string;
    moduleName?: string;
    sourcePath?: string;
    targetPath?: string;
}

/**
 * Runs `fn` and records one succeeded/failed write-audit entry around it.
 * The success return may override the audited fields (e.g. the resolved
 * export folder or the post-rename module name).
 */
export async function withWriteAudit<T>(
    target: XlideWriteAuditTarget & { failedSummary: string },
    fn: () => Promise<{ result: T; summary: string } & Partial<XlideWriteAuditTarget>>,
): Promise<{ result: T; summary: string }> {
    const { failedSummary, ...fields } = target;
    try {
        const { result, summary, ...overrides } = await fn();
        recordXlideWriteAudit({
            timestamp: new Date().toISOString(),
            ...fields,
            ...overrides,
            outcome: 'succeeded',
            summary,
        });
        return { result, summary };
    } catch (err) {
        recordXlideWriteAudit({
            timestamp: new Date().toISOString(),
            ...fields,
            outcome: 'failed',
            summary: failedSummary,
            errorCategory: errorCategoryForSupportLog(err),
        });
        throw err;
    }
}

export interface XlideChangeSummary {
    operation: string;
    changed?: readonly string[];
    skipped?: readonly string[];
    removed?: readonly string[];
    failed?: readonly string[];
}

export function formatChangeSummary(summary: XlideChangeSummary): string {
    const parts: string[] = [];
    const changed = summary.changed?.length ?? 0;
    const skipped = summary.skipped?.length ?? 0;
    const removed = summary.removed?.length ?? 0;
    const failed = summary.failed?.length ?? 0;
    parts.push(`${changed} changed`);
    if (skipped > 0) {
        parts.push(`${skipped} skipped`);
    }
    if (removed > 0) {
        parts.push(`${removed} removed`);
    }
    if (failed > 0) {
        parts.push(`${failed} failed`);
    }
    return `${summary.operation}: ${parts.join(', ')}`;
}

export function formatChangeSummaryDetails(summary: XlideChangeSummary): string[] {
    const lines = [formatChangeSummary(summary)];
    appendDetailLines(lines, 'Changed', summary.changed);
    appendDetailLines(lines, 'Skipped', summary.skipped);
    appendDetailLines(lines, 'Removed', summary.removed);
    appendDetailLines(lines, 'Failed', summary.failed);
    return lines;
}

function appendDetailLines(
    lines: string[],
    label: string,
    values: readonly string[] | undefined,
): void {
    if (!values || values.length === 0) {
        return;
    }
    lines.push(`${label}:`);
    for (const value of values) {
        lines.push(`  - ${value}`);
    }
}
