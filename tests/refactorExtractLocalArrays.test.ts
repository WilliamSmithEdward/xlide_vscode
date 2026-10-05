import { describe, expect, it } from 'vitest';
import { extractMethod } from '../src/analyzer/refactor/extractMethod';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { analyzeModule } from '../src/analyzer/diagnostics/analyzeModule';

function fixture(declarations: string, selected: string, eol: string, before = '', after = '', module = '') {
    const prefix = ['Option Explicit', ...(module ? [module] : []), 'Sub Main()', ...(declarations ? [declarations] : []), ...(before ? [before] : []), ''].join(eol);
    const body = selected.replace(/\n/g, eol);
    return { source: prefix + body + eol + (after ? after + eol : '') + 'End Sub' + eol, span: { start: prefix.length, end: prefix.length + body.length }, name: 'Work' };
}
function applied(input: ReturnType<typeof fixture>) {
    const result = extractMethod(input);
    if (!result.ok) throw Error(result.reason);
    const source = applyVbaTextEdits(input.source, result.edits);
    const helper = parseModule(source).members.find(n => n.kind === 'Procedure' && n.name === 'Work');
    if (helper?.kind !== 'Procedure') throw Error('Missing helper');
    expect(analyzeModule(source, { knownIdentifiers: new Set<string>() }).filter(d => d.code === 'undeclared-variable')).toEqual([]);
    return { source, helper, caller: source.slice(0, source.indexOf('Private ')) };
}

describe.each(['\n', '\r\n', '\r'])('Extract Method local arrays with %j', eol => {
    it.each([
        ['Dim x(1 To 2) As Long', 'Debug.Print x(1)', 'ByRef x() As Long'],
        ['Dim x(1 To 2) As Long', 'x(1) = x(1) + 1', 'ByRef x() As Long'],
        ['Dim x() As Long', 'ReDim x(1 To 2)\nx(2) = 7', 'ByRef x() As Long'],
        ['Dim x() As Long', 'ReDim Preserve x(1 To 3)', 'ByRef x() As Long'],
        ['Dim x(1 To 2) As Long', 'Erase x', 'ByRef x() As Long'],
        ['Dim x(1 To 2, 3 To 4) As Long', 'x(2, 4) = 7', 'ByRef x() As Long'],
        ['Dim x(1 To 2) As Variant', 'x(2) = "hello"', 'ByRef x() As Variant'],
        ['Dim x%(1 To 2)', 'x(2) = 7', 'ByRef x%()'],
        ['Dim x(1 To 2)', 'x(2) = 7', 'ByRef x()'],
        ['Dim [Case](1 To 2) As Long', '[Case](2) = 7', 'ByRef [Case]() As Long'],
        ['Dim x(1 To 2) As VBA.Collection', 'Set x(2) = New VBA.Collection', 'ByRef x() As VBA.Collection'],
        ['Dim x(1 To 2) As [Collection]', 'Set x(2) = New Collection', 'ByRef x() As [Collection]'],
    ])('retains the array binding from %s for %s', (decl, selected, binding) => {
        for (const after of ['', selected.includes('[Case]') ? 'Debug.Print [Case](2)' : 'Debug.Print x(2)']) {
            const { source, helper, caller } = applied(fixture(decl, selected, eol, '', after));
            expect(source).toContain('Private Sub Work(' + binding + ')');
            expect(helper.params).toHaveLength(1);
            expect(helper.params[0].isArray).toBe(true);
            expect(helper.params[0].byRef).toBe(true);
            expect(helper.body.some(n => n.kind === 'VariableGroup')).toBe(false);
            expect(caller).toContain(decl);
        }
    });
    it('keeps an array declared inside the selection in the caller', () => {
        const { source, caller } = applied(fixture('', 'Dim x(1 To 2) As Long\nx(2) = 7', eol, '', 'Debug.Print x(2)'));
        expect(caller).toContain('Dim x(1 To 2) As Long' + eol + 'Work x');
        expect(source).toContain('Private Sub Work(ByRef x() As Long)');
        expect(source.slice(source.indexOf('Private Sub Work'))).not.toContain('Dim x');
    });
    it('preserves sibling scalar declarations and a scalar Function result', () => {
        const { source, helper, caller } = applied(fixture('Dim x(1 To 2) As Long, total As Long', 'x(2) = 7\ntotal = x(2) + 1', eol, '', 'Debug.Print x(2), total'));
        expect(caller).toContain('Dim x(1 To 2) As Long, total As Long');
        expect(caller).toContain('total = Work(x)');
        expect(source).toContain('Private Function Work(ByRef x() As Long) As Long');
        expect(helper.body.filter(n => n.kind === 'VariableGroup').flatMap(n => n.declarations).map(n => n.name)).toEqual(['total']);
    });
    it('keeps module DefType inference for an untyped array', () => {
        const { source } = applied(fixture('Dim x(1 To 2)', 'x(2) = 7', eol, '', '', 'DefLng A-Z'));
        expect(source).toContain('Private Sub Work(ByRef x())');
    });
    it('preserves multiple array bindings alongside ordinary parameters and scalar inputs', () => {
        const input = fixture('Dim x(1 To 2) As Long, y() As String, n As Long', 'x(2) = p\nReDim y(1 To 2)\ny(2) = CStr(n)', eol);
        input.source = input.source.replace('Sub Main()', 'Sub Main(ByVal p As Long)');
        input.span.start += 'ByVal p As Long'.length; input.span.end += 'ByVal p As Long'.length;
        const { source } = applied(input);
        expect(source).toContain('Private Sub Work(ByRef p As Long, ByRef x() As Long, ByRef y() As String, ByVal n As Long)');
    });
    it.each([
        ['Dim x(1 To 2) As New Collection', 'x(2).Add 7', 'As New'],
        ['Dim x(1 To 2) As String * 5', 'x(2) = "hello"', 'fixed-length String'],
    ])('refuses a binding the helper cannot preserve: %s', (decl, selected, reason) => {
        const result = extractMethod(fixture(decl, selected, eol));
        expect(result.ok).toBe(false);
        if (result.ok) throw Error('Expected refusal');
        expect(result.reason).toContain(reason);
    });
    it.each(['Dim x(1 To 2) As New Collection', 'Dim x(1 To 2) As String * 5'])('allows unrelated scalar extraction with %s', decl => {
        const { source } = applied(fixture(decl + eol + 'Dim n As Long', 'Debug.Print n', eol));
        expect(source).toContain('Private Sub Work(ByVal n As Long)');
    });
});
