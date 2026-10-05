import { describe, expect, it } from 'vitest';
import { extractMethod } from '../src/analyzer/refactor/extractMethod';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';
function input(declaration: string, body: string, eol: string, after = 'Debug.Print x', before = '', module = '') {
    const convert = (text: string) => text.replace(/\n/g, eol);
    const prefix = convert('Option Explicit\n' + module + '\nSub Main()\n' + declaration + '\n' + before + '\n');
    const selected = convert(body);
    return { source: prefix + selected + eol + convert(after) + eol + 'End Sub' + eol, span: { start: prefix.length, end: prefix.length + selected.length }, name: 'Work' };
}
function applied(i: ReturnType<typeof input>) { const r = extractMethod(i); expect(r.ok).toBe(true); if (!r.ok) throw Error(r.reason); return applyVbaTextEdits(i.source, r.edits); }
describe.each(['\n', '\r\n', '\r'])('Extract Method primitive inputs with %j', eol => {
    it.each(['Byte', 'Integer', 'Long', 'Single', 'Double', 'Currency', 'Date', 'Boolean', 'String'])('preserves indirect writes to %s', type => {
        const output = applied(input('Dim x As ' + type, 'Mutate x', eol, 'Debug.Print x', '', 'Private Sub Mutate(ByRef p As ' + type + ')\np = ' + (type === 'String' ? '"seven"' : '7') + '\nEnd Sub\n'));
        expect(output).toContain('Private Sub Work(ByRef x As ' + type + ')');
        expect(output).toContain('Work x'); expect(output).toContain('Mutate x');
    });
    it('preserves bracketed input spelling in the formal and call argument', () => {
        const output = applied(input('Dim [x] As Long', 'Debug.Print [x]', eol, ''));
        expect(output).toContain('Private Sub Work(ByRef [x] As Long)');
        expect(output).toContain('Work [x]');
    });
    it('retains indirect writes observed inside the selection alone', () => {
        expect(applied(input('Dim x As Long', 'Mutate x\nDebug.Print x', eol, ''))).toContain('Private Sub Work(ByRef x As Long)');
    });
    it('retains the input binding alongside a scalar Function output', () => {
        const output = applied(input('Dim x As Long, result As Long', 'Mutate x\nresult = x + 1', eol, 'Debug.Print x, result'));
        expect(output).toContain('Private Function Work(ByRef x As Long) As Long');
        expect(output).toContain('result = Work(x)');
    });
    it.each(['Debug.Print x', 'x = x & "abcdefgh"'])('refuses fixed-length String input binding for %s', body => {
        const result = extractMethod(input('Dim x As String * 5', body, eol));
        expect(result.ok).toBe(false); if (result.ok) throw Error('Expected refusal');
        expect(result.reason).toContain('fixed-length String');
    });
    it('refuses a fixed-length String among multiple output bindings', () => {
        const result = extractMethod(input('Dim x As String * 5, y As Long', 'x = "abcdefgh"\ny = 7', eol, 'Debug.Print x, y'));
        expect(result.ok).toBe(false); if (result.ok) throw Error('Expected refusal'); expect(result.reason).toContain('fixed-length String');
    });
    it('still preserves a fixed-length String through its local Function output', () => {
        const output = applied(input('Dim x As String * 5', 'x = "abcdefgh"', eol));
        expect(output).toContain('Private Function Work() As String');
        expect(output).toContain('Dim x As String * 5');
    });
    it('allows unrelated primitive extraction beside a fixed-length String', () => {
        const output = applied(input('Dim x As String * 5, n As Long', 'Debug.Print n', eol, ''));
        expect(output).toContain('Private Sub Work(ByRef n As Long)');
    });
});
