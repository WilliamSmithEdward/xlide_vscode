import { describe, expect, it, vi } from 'vitest';
import { AnalysisWorkerState } from '../src/analysisWorkerLogic';
import * as analysis from '../src/vbaModuleAnalysis';

const request = {
    kind: 'analyze' as const, requestId: 1, docKey: 'C', moduleName: 'C', moduleKind: 'class',
    source: 'Option Explicit\nSub Run()\nDebug.Print (1\nEnd Sub\n',
};

describe('completed live diagnostic snapshots', () => {
    it('reuses an identical snapshot without another analyzer pass and retains the new request ID', () => {
        const spy = vi.spyOn(analysis, 'analyzeVbaModuleSource');
        try {
            const state = new AnalysisWorkerState();
            const first = state.handle(request);
            const second = state.handle({ ...request, requestId: 2 });
            expect(spy).toHaveBeenCalledTimes(1);
            expect(second).toEqual({ ...first, requestId: 2 });
        } finally { spy.mockRestore(); }
    });

    it('invalidates on source, cursor, settings, module kind, host, designer, references and sheet changes', () => {
        const changes = [
            { source: request.source.replace('(1', '(1)') },
            { activeIncompleteExpressionOffset: request.source.indexOf('(1') + 2 },
            { severityOverrides: { 'unused-variable': 'off' } },
            { moduleKind: 'standard' }, { moduleName: 'Renamed' }, { host: 'word' },
            { implicitMembers: [{ name: 'Box', type: 'MSForms.TextBox' }] },
            { referencedLibraries: ['VBA'] }, { referencedHosts: ['access'] },
            { designerClass: 'Access.Form' },
            { workbookSheets: [{ name: 'Data', kind: 'worksheet' as const }] },
        ];
        for (const change of changes) {
            const spy = vi.spyOn(analysis, 'analyzeVbaModuleSource');
            try {
                const state = new AnalysisWorkerState();
                state.handle(request);
                const result = state.handle({ ...request, ...change, requestId: 2 });
                expect(spy).toHaveBeenCalledTimes(2);
                expect(result?.kind).toBe('result');
            } finally { spy.mockRestore(); }
        }
    });

    it('clears completed results on forget and on reseeding the owning project', () => {
        const spy = vi.spyOn(analysis, 'analyzeVbaModuleSource');
        try {
            const state = new AnalysisWorkerState();
            const seed = { kind: 'seed' as const, projectKey: 'P', generation: 1,
                modules: [{ moduleName: 'C', source: request.source, type: 'class' }] };
            const seeded = { ...request, projectKey: 'P', generation: 1 };
            state.handle(seed);
            state.handle(seeded);
            state.handle({ kind: 'forget', docKey: request.docKey });
            state.handle(seeded);
            state.handle(seed);
            state.handle(seeded);
            expect(spy).toHaveBeenCalledTimes(3);
        } finally { spy.mockRestore(); }
    });

    it('keeps a completed standalone snapshot when another project is seeded', () => {
        const spy = vi.spyOn(analysis, 'analyzeVbaModuleSource');
        try {
            const state = new AnalysisWorkerState();
            state.handle(request);
            state.handle({ kind: 'seed', projectKey: 'Other', generation: 1, modules: [] });
            state.handle(request);
            expect(spy).toHaveBeenCalledTimes(1);
        } finally { spy.mockRestore(); }
    });

    it('does not cache a recovered analysis failure', () => {
        const spy = vi.spyOn(analysis, 'analyzeVbaModuleSource').mockReturnValue({
            diagnostics: [], suppressedDiagnostics: [], suppressedCount: 0,
            analysisFailures: [{ stage: 'structural', message: 'test failure' }],
        });
        try {
            const state = new AnalysisWorkerState();
            state.handle(request);
            state.handle(request);
            expect(spy).toHaveBeenCalledTimes(2);
        } finally { spy.mockRestore(); }
    });
});
