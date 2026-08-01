import { describe, expect, it, beforeEach } from 'vitest';
import {
    clearPerformanceTrace,
    formatPerformanceSnapshot,
    isCancellationLike,
    measurePerformance,
    measurePerformanceSync,
    recentPerformanceTraceEntries,
    recordPerformanceTrace,
} from '../src/performanceTrace';

describe('performanceTrace', () => {
    beforeEach(() => {
        clearPerformanceTrace();
    });

    it('keeps a bounded recent event history', () => {
        for (let i = 0; i < 505; i++) {
            recordPerformanceTrace({
                name: 'completion',
                startedAt: `2026-01-01T00:00:${String(i % 60).padStart(2, '0')}.000Z`,
                durationMs: i,
                outcome: 'ok',
            });
        }

        const entries = recentPerformanceTraceEntries(600);

        expect(entries).toHaveLength(500);
        expect(entries[0].durationMs).toBe(5);
        expect(entries[499].durationMs).toBe(504);
    });

    it('formats recent events and grouped summary rows', () => {
        recordPerformanceTrace({
            name: 'completion',
            startedAt: '2026-01-01T00:00:00.000Z',
            durationMs: 10,
            outcome: 'ok',
        });
        recordPerformanceTrace({
            name: 'completion',
            startedAt: '2026-01-01T00:00:01.000Z',
            durationMs: 30,
            outcome: 'ok',
        });

        const snapshot = formatPerformanceSnapshot();

        expect(snapshot).toContain('XLIDE Performance Snapshot');
        expect(snapshot).toContain('[perf] 2026-01-01T00:00:00.000Z completion ok 10 ms');
        expect(snapshot).toContain('completion: count=2, avg=20.0 ms, max=30 ms');
    });

    it('records failed async measurements and rethrows', async () => {
        await expect(measurePerformance('workbookEngine.call', 'readModule', async () => {
            throw new Error('boom');
        })).rejects.toThrow('boom');

        const [entry] = recentPerformanceTraceEntries();
        expect(entry.name).toBe('workbookEngine.call');
        expect(entry.detail).toBe('readModule');
        expect(entry.outcome).toBe('failed');
    });

    it('records canceled sync measurements and rethrows', () => {
        expect(() => measurePerformanceSync('completion', 'member', () => {
            const err = new Error('Canceled by newer request');
            err.name = 'CancellationError';
            throw err;
        })).toThrow('Canceled by newer request');

        const [entry] = recentPerformanceTraceEntries();
        expect(entry.name).toBe('completion');
        expect(entry.detail).toBe('member');
        expect(entry.outcome).toBe('canceled');
        expect(isCancellationLike({ name: 'CancellationError', message: '' })).toBe(true);
    });
});
