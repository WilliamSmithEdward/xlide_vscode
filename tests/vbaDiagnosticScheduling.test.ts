import { describe, expect, it, vi } from 'vitest';
import type * as vscode from 'vscode';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());

import { editScheduleDelaysFor } from '../src/vbaLiveDiagnostics';

// editScheduleDelaysFor reads only lineCount and uri.scheme.
function doc(lineCount: number, scheme: string): vscode.TextDocument {
    return { lineCount, uri: { scheme } } as unknown as vscode.TextDocument;
}

const XLIDE = 'xlide-vba';

describe('diagnostic pass scheduling', () => {
    it('keeps the flat fast cadence when the worker is healthy, at any size', () => {
        // Both passes run off-thread, so a 24k-line class types like a small
        // module: local at 90ms, full at 450ms, no host contention to pace.
        for (const lines of [200, 8000, 24000]) {
            expect(editScheduleDelaysFor(doc(lines, XLIDE), true)).toEqual({
                localDelayMs: 90,
                fullDelayMs: 450,
            });
        }
    });

    it('keeps small modules fast even in-host', () => {
        expect(editScheduleDelaysFor(doc(500, XLIDE), false)).toEqual({
            localDelayMs: 90,
            fullDelayMs: 450,
        });
    });

    it('drops the in-host local pass for large workbook modules', () => {
        // The in-host local pass is a full module analysis (~700ms on a
        // 24k-line class). With no worker, the paced full pass covers the
        // document; running local too would block the host twice per pause.
        const delays = editScheduleDelaysFor(doc(24000, XLIDE), false);
        expect(delays.localDelayMs).toBeUndefined();
        expect(delays.fullDelayMs).toBeGreaterThan(450);
        expect(delays.fullDelayMs).toBeLessThanOrEqual(2000);
    });

    it('paces (but keeps) the local pass for large loose .bas files', () => {
        // Loose files never get a full pass, so local is their only analysis:
        // it must survive, just far behind the typing burst.
        const delays = editScheduleDelaysFor(doc(24000, 'file'), false);
        expect(delays.localDelayMs).toBe(delays.fullDelayMs);
        expect(delays.localDelayMs).toBeGreaterThan(450);
    });

    it('scales the in-host backoff with module size and caps at 2s', () => {
        const at10k = editScheduleDelaysFor(doc(10000, XLIDE), false).fullDelayMs ?? 0;
        const at24k = editScheduleDelaysFor(doc(24000, XLIDE), false).fullDelayMs ?? 0;
        expect(at10k).toBeGreaterThan(450);
        expect(at24k).toBeGreaterThanOrEqual(at10k);
        expect(at24k).toBeLessThanOrEqual(2000);
    });
});
