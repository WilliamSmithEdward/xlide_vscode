export type XlideWriteAuditOutcome = 'succeeded' | 'failed' | 'skipped';

export interface XlideWriteAuditEntry {
    timestamp: string;
    command: string;
    operation: string;
    outcome: XlideWriteAuditOutcome;
    workbookPath?: string;
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
    return writeAudit.slice(-Math.max(0, limit));
}

export function clearXlideWriteAudit(): void {
    writeAudit.length = 0;
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
