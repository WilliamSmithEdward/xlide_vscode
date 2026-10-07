import { describe, expect, it } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

describe('deleted-object locations on large modules', () => {
    for (const eol of ['\n', '\r\n', '\r']) {
        it(`reports the delete location with ${JSON.stringify(eol)} lines`, () => {
            const source = [
                'Option Explicit',
                ...Array.from({ length: 1000 }, () => "' padding before the procedure"),
                'Function F() As String',
                'Dim n As Name',
                'Set n = ThisWorkbook.Names.Add("xlideName", "=1")',
                'n.Delete',
                'F = n.Name',
                'End Function',
            ].join(eol);
            const findings = analyzeModule(source).filter(d => d.code === 'object-used-after-delete');
            expect(findings).toHaveLength(1);
            expect(findings[0].message).toContain('deleted on line 1005');
            expect(source.slice(findings[0].span.start, findings[0].span.end)).toBe('n');
        });
    }
});
