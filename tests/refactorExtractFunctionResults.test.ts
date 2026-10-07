import { describe, expect, it } from 'vitest';
import { extractMethod } from '../src/analyzer/refactor/extractMethod';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { analyzeModule } from '../src/analyzer/diagnostics/analyzeModule';
function fixture(header: string, selected: string, eol: string, before = '', after = '', module = '', name = 'Work') {
    const convert = (s: string) => s.replace(/\n/g, eol);
    const prefix = convert('Option Explicit\n' + (module ? module + '\n' : '') + header + '\n' + (before ? before + '\n' : ''));
    const body = convert(selected);
    return { source: prefix + body + eol + (after ? convert(after) + eol : '') + 'End ' + (header.includes('Property') ? 'Property' : 'Function') + eol,
        span: { start: prefix.length, end: prefix.length + body.length }, name };
}
function applied(input: ReturnType<typeof fixture>) {
    const result = extractMethod(input);
    if (!result.ok) throw Error(result.reason);
    const source = applyVbaTextEdits(input.source, result.edits);
    expect(analyzeModule(source, { knownIdentifiers: new Set<string>() }).filter(d => d.code === 'undeclared-variable')).toEqual([]);
    const helper = parseModule(source).members.find(n => n.kind === 'Procedure' && n.name === input.name);
    if (helper?.kind !== 'Procedure') throw Error('Missing helper');
    return { source, helper, body: source.slice(source.indexOf('Private ')) };
}
describe.each(['\n', '\r\n', '\r'])('Extract Method implicit results with %j', eol => {
    it.each(['Compute = 7', 'Compute = Compute + 1', 'Mutate Compute'])('forwards the original result for %s', selected => {
        const module = 'Private Sub Mutate(ByRef n As Long)\nn = 7\nEnd Sub';
        const { source, helper } = applied(fixture('Public Function Compute() As Long', selected, eol, '', '', module));
        expect(source).toContain('Private Sub Work(ByRef ComputeResult As Long)');
        expect(source).toContain(eol + 'Work Compute' + eol);
        expect(helper.params.map(p => p.name)).toEqual(['ComputeResult']);
    });
    it.each([
        ['Public Function Compute() As String', 'Compute = "hello"', 'ByRef ComputeResult As String'],
        ['Public Function Compute() As Collection', 'Set Compute = New Collection\nCompute.Add 7', 'ByRef ComputeResult As Collection'],
        ['Public Function Compute%()', 'Compute = 7', 'ByRef ComputeResult%'],
        ['Public Function Compute$()', 'Compute = "hello"', 'ByRef ComputeResult$'],
        ['Public Function Compute()', 'Compute = 7', 'ByRef ComputeResult'],
        ['Public Property Get Compute() As Long', 'Compute = 7', 'ByRef ComputeResult As Long'],
        ['Public Function Compute() As VBA.Collection', 'Set Compute = New VBA.Collection', 'ByRef ComputeResult As VBA.Collection'],
        ['Public Function Compute() As [Collection]', 'Set Compute = New Collection', 'ByRef ComputeResult As [Collection]'],
        ['Public Function Compute( _\nByVal n As Long _\n) As _\nLong', 'Compute = n', 'ByRef ComputeResult As Long, ByRef n As Long'],
    ])('preserves the result declaration from %s', (header, selected, binding) => {
        const { source } = applied(fixture(header, selected, eol));
        expect(source).toContain('Private Sub Work(' + binding + ')');
        expect(source).toContain('ComputeResult');
    });
    it('keeps recursive calls separate from result writes', () => {
        const { source } = applied(fixture('Public Function Compute(ByVal n As Long) As Long', 'If n > 0 Then\nCompute = n + Compute(n - 1)\nEnd If', eol));
        expect(source).toContain('ComputeResult = n + Compute(n - 1)');
        expect(source).toContain('Work Compute, n');
    });
    it('keeps explicit zero-argument recursive calls separate from result reads', () => {
        const { source } = applied(fixture('Public Function Compute() As Long', 'Compute = Compute + Compute()', eol));
        expect(source).toContain('ComputeResult = ComputeResult + Compute()');
    });
    it('preserves a result object receiver even without a selected Set assignment', () => {
        const { source } = applied(fixture('Public Function Compute() As Collection', 'Compute.Add 7', eol, 'Set Compute = New Collection'));
        expect(source).toContain('Private Sub Work(ByRef ComputeResult As Collection)');
        expect(source).toContain('ComputeResult.Add 7');
    });
    it('preserves bang member receivers without a selected Set assignment', () => {
        const { source } = applied(fixture('Public Function Compute() As Collection', 'Compute!child.Add 7', eol, 'Set Compute = New Collection'));
        expect(source).toContain('ComputeResult!child.Add 7');
        expect(source).toContain('Private Sub Work(ByRef ComputeResult As Collection)');
    });
    it('does not create a result binding for qualified member names alone', () => {
        const { source } = applied(fixture('Public Function Compute() As Long', 'Debug.Print obj.Compute', eol, 'Dim obj As Object'));
        expect(source).toContain('Private Sub Work(ByRef obj As Object)');
        expect(source).not.toContain('ComputeResult');
    });
    it('retains selective DefType by keeping the original first letter', () => {
        const { source } = applied(fixture('Public Function Compute()', 'Compute = 7', eol, '', '', 'DefLng C-C'));
        expect(source).toContain('Private Sub Work(ByRef ComputeResult)');
    });
    it('does not rewrite strings or comments with the same spelling', () => {
        const { source } = applied(fixture('Public Function Compute() As Long', 'Compute = 7 \' Compute\nDebug.Print "Compute"', eol));
        expect(source).toContain('ComputeResult = 7');
        expect(source).toContain('"Compute"'); expect(source).toContain("' Compute");
    });
    it('does not shadow existing source names or the chosen helper name', () => {
        const { source } = applied(fixture('Public Function Compute() As Long', 'Compute = 7', eol, 'Dim ComputeResult As Long', '', '', 'ComputeResult2'));
        expect(source).toContain('Private Sub ComputeResult2(ByRef ComputeResult3 As Long)');
        expect(source).toContain('ComputeResult3 = 7');
    });
    it('bounds the generated result name for a long original identifier', () => {
        const name = 'C' + 'a'.repeat(249);
        const { source, helper } = applied(fixture('Public Function ' + name + '() As Long', name + ' = 7', eol));
        expect(helper.params[0].name).toBe(name.slice(0, 120) + 'Result');
        expect(helper.params[0].name.length).toBeLessThanOrEqual(255);
        expect(source).toContain('Work ' + name);
    });
    it('retains scalar Function outputs alongside the original result binding', () => {
        const { source, helper } = applied(fixture('Public Function Compute() As Long', 'Compute = 7\nn = Compute + 1', eol, 'Dim n As Long', 'Compute = Compute + n'));
        expect(helper.procKind).toBe('Function');
        expect(source).toContain('Private Function Work(ByRef ComputeResult As Long) As Long');
        expect(source).toContain('n = Work(Compute)');
        expect(source).toContain('n = ComputeResult + 1');
    });
    it('leaves qualified references and type names untouched', () => {
        const { source } = applied(fixture('Public Function Compute() As Long', 'Compute = 7\nDebug.Print obj.Compute\nDebug.Print TypeOf obj Is Compute', eol, 'Dim obj As Object'));
        expect(source).toContain('ComputeResult = 7');
        expect(source).toContain('obj.Compute'); expect(source).toContain('Is Compute');
    });
    it('does not create a result binding for a selection containing only recursive calls', () => {
        const { source } = applied(fixture('Public Function Compute(ByVal n As Long) As Long', 'n = Compute(n - 1)', eol, '', 'Compute = n'));
        expect(source).toContain('Private Sub Work(ByRef n As Long)');
        expect(source).not.toContain('ComputeResult');
    });
    it.each([
        ['Public Function Compute() As Long()', 'Compute = x', 'array-valued', 'Dim x(1 To 2) As Long'],
        ['Public Function Compute() As Long', 'Compute = 7\nExit Function', 'exits the original', ''],
        ['Public Property Get Compute() As Long', 'Compute = 7\nExit Property', 'exits the original', ''],
    ])('refuses an unsupported result transfer from %s', (header, selected, reason, before) => {
        const result = extractMethod(fixture(header, selected, eol, before));
        expect(result.ok).toBe(false);
        if (result.ok) throw Error('Expected refusal');
        expect(result.reason).toContain(reason);
    });
});
