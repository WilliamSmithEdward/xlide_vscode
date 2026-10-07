import { describe, expect, it, vi } from 'vitest';
import { AnalysisWorkerState } from '../src/analysisWorkerLogic';
import { DIAGNOSTIC_RULE_REGISTRY } from '../src/analyzer/diagnostics/registry';

const request = { kind: 'analyze' as const, requestId: 1, docKey: 'M', moduleName: 'M',
    source: 'Sub P()\nDim n As Long\nDebug.Print n\nEnd Sub' };

describe('cooperative analysis cancellation', () => {
    it('rejects an already cancelled request before reading completed results', () => {
        const state = new AnalysisWorkerState();
        const initial = state.handle(request);
        const signal = new Int32Array(new SharedArrayBuffer(4));
        Atomics.store(signal, 0, 1);
        expect(state.handle({ ...request, requestId: 2, cancellationSignal: signal }))
            .toEqual({ kind: 'cancelled', requestId: 2, docKey: 'M' });
        expect(state.handle({ ...request, requestId: 3 })).toEqual({ ...initial, requestId: 3 });
    });

    it('stops within a rule without publishing or caching partial diagnostics', () => {
        const state = new AnalysisWorkerState();
        const initial = state.handle(request);
        const signal = new Int32Array(new SharedArrayBuffer(4));
        const rule = DIAGNOSTIC_RULE_REGISTRY.find(rule => rule.name === 'overflow')!;
        const run = rule.run!;
        const spy = vi.spyOn(rule, 'run').mockImplementation((ctx, push) => {
            Atomics.store(signal, 0, 1);
            return run(ctx, push);
        });
        const changed = { ...request, requestId: 2, source: request.source.replace('Debug.Print n', 'n = 1\nDebug.Print n') };
        try {
            expect(state.handle({ ...changed, cancellationSignal: signal }))
                .toEqual({ kind: 'cancelled', requestId: 2, docKey: 'M' });
            expect(spy).toHaveBeenCalledOnce();
        } finally { spy.mockRestore(); }
        expect(state.handle({ ...request, requestId: 3 })).toEqual({ ...initial, requestId: 3 });
        const result = state.handle({ ...changed, requestId: 4 });
        const full = new AnalysisWorkerState().handle(changed);
        expect(result?.kind).toBe('result');
        expect(full?.kind).toBe('result');
        if (result?.kind === 'result' && full?.kind === 'result') {
            expect(result.diagnostics).toEqual(full.diagnostics);
            expect(result.suppressedDiagnostics).toEqual(full.suppressedDiagnostics);
            expect(result.analysisFailures).toBeUndefined();
        }
    });

    it('matches full diagnostics with cancellation checks enabled but inactive', () => {
        const signal = new Int32Array(new SharedArrayBuffer(4));
        const state = new AnalysisWorkerState();
        expect(state.handle({ ...request, cancellationSignal: signal })).toEqual(new AnalysisWorkerState().handle(request));
    });
});
