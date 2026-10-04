import { describe, expect, it } from 'vitest';
import { extractMethod } from '../src/analyzer/refactor/extractMethod';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';
const procedures = [
    ['Sub Main()', 'End Sub'],
    ['Function Main() As Long', 'End Function'],
    ['Property Get Main() As Long', 'End Property'],
] as const;
const endings = [['LF', '\n'], ['CRLF', '\r\n'], ['CR', '\r']] as const;
function fixture(header: string, closer: string, eol: string, trailing: boolean) {
    const prefix = ['Option Explicit', header, '    Dim one As Long', ''].join(eol);
    const selected = '    Debug.Print one';
    const source = prefix + selected + eol + closer + (trailing ? eol : '');
    return { source, span: { start: prefix.length, end: prefix.length + selected.length }, name: 'Work' };
}
for (const [ending, eol] of endings) {
    describe('last statement with ' + ending, () => {
        for (const [header, closer] of procedures) {
            it.each([false, true])('extracts only the last body statement in ' + header + ' with trailing newline=%s', trailing => {
                const input = fixture(header, closer, eol, trailing);
                const result = extractMethod(input);
                expect(result.ok).toBe(true);
                if (!result.ok) { throw new Error(result.reason); }
                const actual = applyVbaTextEdits(input.source, result.edits);
                expect(actual).toContain(header + eol + '    Dim one As Long' + eol + '    Work one' + eol + closer);
                expect(actual).toContain('Private Sub Work(ByVal one As Long)' + eol + '    Debug.Print one' + eol + 'End Sub');
                expect(input.source.slice(input.span.start, input.span.end)).toBe('    Debug.Print one');
            });

            it.each([false, true])('moves the last statement and its declaration in ' + header + ' with trailing newline=%s', trailing => {
                const base = fixture(header, closer, eol, trailing);
                const prefix = base.source.slice(0, base.span.start);
                const selected = '    one = 1';
                const input = { ...base, source: prefix + selected + eol + closer + (trailing ? eol : ''), span: { start: prefix.length, end: prefix.length + selected.length } };
                const result = extractMethod(input);
                expect(result.ok).toBe(true);
                if (!result.ok) { throw new Error(result.reason); }
                const actual = applyVbaTextEdits(input.source, result.edits);
                expect(actual).toContain('Option Explicit' + eol + header + eol + '    Work' + eol + closer);
                expect(actual).toContain('Private Sub Work()' + eol + '    Dim one As Long' + eol + selected + eol + 'End Sub');
                expect(actual.split('Dim one As Long')).toHaveLength(2);
            });

            it('still refuses the actual header, End line and partial statement in ' + header, () => {
                const input = fixture(header, closer, eol, true);
                const headerResult = extractMethod({ ...input, span: { start: input.source.indexOf(header), end: input.span.end } });
                expect(headerResult.ok).toBe(false);
                if (headerResult.ok) { throw new Error('Accepted header'); }
                expect(headerResult.reason).toMatch(/header or End line|whole statements/);
                const endResult = extractMethod({ ...input, span: { start: input.span.start, end: input.source.indexOf(closer) + closer.length } });
                expect(endResult.ok).toBe(false);
                if (endResult.ok) { throw new Error('Accepted End line'); }
                expect(endResult.reason).toMatch(/header or End line/);
                const partialResult = extractMethod({ ...input, span: { start: input.span.start + 5, end: input.span.end } });
                expect(partialResult.ok).toBe(false);
                if (partialResult.ok) { throw new Error('Accepted partial statement'); }
                expect(partialResult.reason).toMatch(/whole statements|header or End line/);
            });
        }
    });
}
