interface PerformanceTraceEntry {
    name: string;
    startedAt: string;
    durationMs: number;
    detail?: string;
    outcome?: 'ok' | 'failed' | 'canceled' | 'superseded';
}

interface ActivePerformanceTrace {
    end(outcome?: PerformanceTraceEntry['outcome'], detail?: string): void;
}

const MAX_PERFORMANCE_TRACE_ENTRIES = 500;
const SLOW_TRACE_THRESHOLD_MS = 100;
const performanceTraceEntries: PerformanceTraceEntry[] = [];
let traceLogger: ((line: string) => void) | undefined;
let traceLoggingEnabled: (() => boolean) | undefined;

function startPerformanceTrace(name: string, detail?: string): ActivePerformanceTrace {
    const startedAtMs = Date.now();
    const startedAt = new Date(startedAtMs).toISOString();
    let ended = false;
    return {
        end(outcome = 'ok', endDetail = detail): void {
            if (ended) {
                return;
            }
            ended = true;
            recordPerformanceTrace({
                name,
                startedAt,
                durationMs: Date.now() - startedAtMs,
                detail: endDetail,
                outcome,
            });
        },
    };
}

async function measurePerformance<T>(
    name: string,
    detail: string | undefined,
    work: () => Promise<T>,
): Promise<T> {
    const trace = startPerformanceTrace(name, detail);
    try {
        const result = await work();
        trace.end('ok');
        return result;
    } catch (err) {
        trace.end(isCancellationLike(err) ? 'canceled' : 'failed');
        throw err;
    }
}

function measurePerformanceSync<T>(
    name: string,
    detail: string | undefined,
    work: () => T,
): T {
    const trace = startPerformanceTrace(name, detail);
    try {
        const result = work();
        trace.end('ok');
        return result;
    } catch (err) {
        trace.end(isCancellationLike(err) ? 'canceled' : 'failed');
        throw err;
    }
}

function recordPerformanceTrace(entry: PerformanceTraceEntry): void {
    performanceTraceEntries.push(entry);
    if (performanceTraceEntries.length > MAX_PERFORMANCE_TRACE_ENTRIES) {
        performanceTraceEntries.splice(0, performanceTraceEntries.length - MAX_PERFORMANCE_TRACE_ENTRIES);
    }
    if (traceLoggingEnabled?.() && entry.durationMs >= SLOW_TRACE_THRESHOLD_MS) {
        traceLogger?.(formatTraceEntry(entry));
    }
}

function recentPerformanceTraceEntries(limit = MAX_PERFORMANCE_TRACE_ENTRIES): PerformanceTraceEntry[] {
    return performanceTraceEntries.slice(-Math.max(0, limit));
}

function clearPerformanceTrace(): void {
    performanceTraceEntries.length = 0;
}

function setPerformanceTraceLogger(
    logger: ((line: string) => void) | undefined,
    enabled: (() => boolean) | undefined,
): void {
    traceLogger = logger;
    traceLoggingEnabled = enabled;
}

function formatPerformanceSnapshot(limit = 120): string {
    const entries = recentPerformanceTraceEntries(limit);
    const lines = [
        'XLIDE Performance Snapshot',
        `Generated: ${new Date().toISOString()}`,
        `Entries: ${entries.length}`,
        '',
        'Recent Events',
    ];
    if (entries.length === 0) {
        lines.push('(none)');
    } else {
        for (const entry of entries) {
            lines.push(formatTraceEntry(entry));
        }
    }
    lines.push('', 'Summary');
    const grouped = summarizePerformanceEntries(entries);
    if (grouped.length === 0) {
        lines.push('(none)');
    } else {
        for (const item of grouped) {
            lines.push(
                `${item.name}: count=${item.count}, avg=${item.averageMs.toFixed(1)} ms, max=${item.maxMs} ms`,
            );
        }
    }
    return `${lines.join('\n')}\n`;
}

function summarizePerformanceEntries(entries: readonly PerformanceTraceEntry[]): Array<{
    name: string;
    count: number;
    averageMs: number;
    maxMs: number;
}> {
    const grouped = new Map<string, { count: number; totalMs: number; maxMs: number }>();
    for (const entry of entries) {
        const current = grouped.get(entry.name) ?? { count: 0, totalMs: 0, maxMs: 0 };
        current.count += 1;
        current.totalMs += entry.durationMs;
        current.maxMs = Math.max(current.maxMs, entry.durationMs);
        grouped.set(entry.name, current);
    }
    return [...grouped.entries()]
        .map(([name, value]) => ({
            name,
            count: value.count,
            averageMs: value.totalMs / value.count,
            maxMs: value.maxMs,
        }))
        .sort((a, b) => b.maxMs - a.maxMs || a.name.localeCompare(b.name));
}

function formatTraceEntry(entry: PerformanceTraceEntry): string {
    const outcome = entry.outcome ? ` ${entry.outcome}` : '';
    const detail = entry.detail ? ` ${entry.detail}` : '';
    return `[perf] ${entry.startedAt} ${entry.name}${outcome} ${entry.durationMs} ms${detail}`;
}

function isCancellationLike(err: unknown): boolean {
    if (!err || typeof err !== 'object') {
        return false;
    }
    const name = 'name' in err ? String((err as { name?: unknown }).name) : '';
    const message = 'message' in err ? String((err as { message?: unknown }).message) : '';
    return /cancel/i.test(`${name}\n${message}`);
}

export {
    clearPerformanceTrace,
    formatPerformanceSnapshot,
    isCancellationLike,
    measurePerformance,
    measurePerformanceSync,
    recentPerformanceTraceEntries,
    recordPerformanceTrace,
    setPerformanceTraceLogger,
    startPerformanceTrace,
    type PerformanceTraceEntry,
};
