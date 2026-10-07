import { describe, expect, it } from 'vitest';
import { analyzeVbaModuleSource } from '../src/vbaModuleAnalysis';
import { AnalysisWorkerState } from '../src/analysisWorkerLogic';

describe('early error analysis', () => {
    const cases = [
        'MsgBox "hello world"-', 'Debug.Print (1', 'MsgBox "hello _',
        'If True Then\nDebug.Print 1', 'Debug.Print 1 + * 2',
        '#If False Then\nMsgBox "hello world"-\n#End If',
        "' @Ignore invalid-expression-syntax\nMsgBox \"hello world\"-",
        'On Error Resume Next\nMsgBox "hello world"-',
        'If False Then\nMsgBox "hello world"-\nEnd If',
        'Debug.Print xlideMissingIdentifier', 'xlideMissingProcedure',
        'Debug.Print Left$("abc")', 'Dim ws As Worksheet\nws.XlideMissingMember = 1',
        'Dim n As Long\nn = "abc"', 'Debug.Print CByte(256)',
        'Dim a(1) As Long\nDebug.Print a(2)', 'Dim c As Collection\nDebug.Print c.Count',
        'Dim s As String * 0', 'Debug.Print 1 / 0',
    ];
    it.each(cases)('matches full-analysis error findings for %s', body => {
        const source = `Option Explicit\nSub Run()\n${body}\nEnd Sub\n`;
        const context = { knownIdentifiers: new Set<string>(), knownProcedures: new Set<string>() };
        const full = analyzeVbaModuleSource({ source, ...context });
        const early = analyzeVbaModuleSource({ source, errorsOnly: true, ...context });
        const select = (findings: typeof full.diagnostics) => findings.filter(d => d.severity === 'error')
            .sort((a, b) => (a.code ?? '').localeCompare(b.code ?? '') || a.span.start - b.span.start);
        expect(select(early.diagnostics)).toEqual(select(full.diagnostics));
        expect(select(early.suppressedDiagnostics)).toEqual(select(full.suppressedDiagnostics));
        expect(early.analysisFailures).toBeUndefined();
    });

    it('reports the exact MsgBox trailing-minus hard error in the early pass', () => {
        const source = 'Sub Run()\nMsgBox "hello world"-\nEnd Sub\n';
        expect(analyzeVbaModuleSource({ source, errorsOnly: true }).diagnostics
            .some(d => d.code === 'string-arithmetic-coercion')).toBe(true);
    });

    it('retains incomplete syntax for the editor to hold and reveal on line exit', () => {
        const source = 'Sub Run()\nDebug.Print (1\nEnd Sub\n';
        expect(analyzeVbaModuleSource({ source, errorsOnly: true }).diagnostics
            .some(d => d.code === 'unbalanced-parens')).toBe(true);
        expect(analyzeVbaModuleSource({ source, activeIncompleteExpressionOffset: source.indexOf('(1') + 2 }).diagnostics
            .some(d => d.code === 'unbalanced-parens')).toBe(false);
    });

    it('keeps error-only results out of complete-result and incremental semantic reuse', () => {
        const state = new AnalysisWorkerState();
        const request = { kind: 'analyze' as const, requestId: 1, docKey: 'D', moduleName: 'C', moduleKind: 'class',
            source: 'Option Explicit\nSub Run()\nDim x As Long\nMsgBox "hello world"-\nEnd Sub\n' };
        const early = state.handle({ ...request, errorsOnly: true });
        const full = state.handle({ ...request, requestId: 2 });
        expect(early?.kind).toBe('result');
        expect(full?.kind).toBe('result');
        if (early?.kind !== 'result' || full?.kind !== 'result') { return; }
        expect(early.incrementalMode).toBe('full');
        expect(full.incrementalMode).toBe('full');
        expect(early.diagnostics.some(d => d.code === 'unused-variable')).toBe(false);
        expect(full.diagnostics.some(d => d.code === 'unused-variable')).toBe(true);
    });
    it.each([
        ['Debug.Print xlideMissingIdentifier', 'undeclared-variable'],
        ['xlideMissingProcedure', 'unknown-call'],
        ['Debug.Print Left$("abc")', 'argument-count'],
        ['Dim ws As Worksheet\nws.XlideMissingMember = 1', 'member-not-found'],
        ['Dim n As Long\nn = "abc"', 'assignment-type-mismatch'],
        ['Dim a(1) As Long\nDebug.Print a(2)', 'array-subscript-out-of-bounds'],
        ['Dim c As Collection\nDebug.Print c.Count', 'object-variable-not-set'],
    ])('reports semantic error %s through the early pass', (body, code) => {
        const result = analyzeVbaModuleSource({ source: `Option Explicit\nSub P()\n${body}\nEnd Sub`,
            errorsOnly: true, knownIdentifiers: new Set(), knownProcedures: new Set() });
        expect(result.diagnostics.some(d => d.code === code && d.severity === 'error')).toBe(true);
        expect(result.diagnostics.every(d => d.severity === 'error')).toBe(true);
    });

    it('honors severity downgrades in the error pass', () => {
        const source = 'Option Explicit\nSub P()\nDebug.Print missingVariable\nEnd Sub';
        const context = { source, knownIdentifiers: new Set<string>(),
            severityOverrides: { 'undeclared-variable': 'warning' as const } };
        expect(analyzeVbaModuleSource(context).diagnostics.some(d => d.code === 'undeclared-variable' && d.severity === 'warning')).toBe(true);
        expect(analyzeVbaModuleSource({ ...context, errorsOnly: true }).diagnostics.some(d => d.code === 'undeclared-variable')).toBe(false);
    });

});
