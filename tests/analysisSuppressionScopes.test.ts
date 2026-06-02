import { describe, expect, it } from 'vitest';
import { validAnalysisSuppressionScopesForDiagnostic } from '../src/analysisSuppressionScopes';
import { analyzeVbaModuleSource } from '../src/vbaModuleAnalysis';

describe('validAnalysisSuppressionScopesForDiagnostic', () => {
    it('limits module-level Option Explicit warnings to module suppression', () => {
        const source = 'Sub T()\nEnd Sub\n';
        const diagnostic = diagnosticByCode(source, 'option-explicit-missing');

        expect(validAnalysisSuppressionScopesForDiagnostic(
            source,
            diagnostic.code,
            diagnostic.span.start,
        )).toEqual(['module']);
    });

    it('allows member and module suppression for procedure-body diagnostics outside nested blocks', () => {
        const source =
            'Sub T()\n' +
            '    Const MAX As Long = 10\n' +
            '    MAX = 1\n' +
            'End Sub\n';
        const diagnostic = diagnosticByCode(source, 'const-assignment');

        expect(validAnalysisSuppressionScopesForDiagnostic(
            source,
            diagnostic.code,
            diagnostic.span.start,
        )).toEqual(['member', 'module']);
    });

    it('allows block suppression for diagnostics inside executable blocks', () => {
        const source =
            'Sub T()\n' +
            '    Const MAX As Long = 10\n' +
            '    If True Then\n' +
            '        MAX = 1\n' +
            '    End If\n' +
            'End Sub\n';
        const diagnostic = diagnosticByCode(source, 'const-assignment');

        expect(validAnalysisSuppressionScopesForDiagnostic(
            source,
            diagnostic.code,
            diagnostic.span.start,
        )).toEqual(['block', 'member', 'module']);
    });
});

function diagnosticByCode(source: string, code: string) {
    const diagnostic = analyzeVbaModuleSource({ source }).diagnostics.find((item) => item.code === code);
    if (!diagnostic) {
        throw new Error(`Expected diagnostic ${code}`);
    }
    return diagnostic;
}
