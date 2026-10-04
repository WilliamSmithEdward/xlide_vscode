import { describe, expect, it } from 'vitest';
import { extractMethod } from '../src/analyzer/refactor/extractMethod';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';
import { analyzeModule } from '../src/analyzer/diagnostics/analyzeModule';
import { parseModule } from '../src/analyzer/parser/parseModule';

function extract(header: string, body: string, eol: string, before = '', after = '', declarations = '') {
    const prefix = ['Option Explicit', ...(before ? [before] : []), header, ...(declarations ? [declarations] : []), ''].join(eol);
    const source = prefix + body + eol + after + 'End Sub' + eol;
    const result = extractMethod({ source, span: { start: prefix.length, end: prefix.length + body.length }, name: 'Work' });
    if (!result.ok) throw Error(result.reason);
    const applied = applyVbaTextEdits(source, result.edits);
    expect(analyzeModule(applied, { knownIdentifiers: new Set<string>() }).filter(d => d.code === 'undeclared-variable')).toEqual([]);
    const helper = parseModule(applied).members.find(n => n.kind === 'Procedure' && n.name === 'Work');
    if (helper?.kind !== 'Procedure') throw Error('Missing helper');
    return { applied, helper };
}

describe.each(['\n', '\r\n', '\r'])('Extract Method parameter bindings with %j', eol => {
    it.each(['ByVal x As Long', 'ByRef x As Long', 'x As Long'])('retains the procedure variable for %s', parameter => {
        for (const body of ['x = 1', 'x = x + 1', 'Debug.Print x']) {
            for (const after of ['', 'Debug.Print x' + eol]) {
                const { applied, helper } = extract('Sub Main(' + parameter + ')', body, eol, '', after);
                expect(applied).toContain('Private Sub Work(ByRef x As Long)');
                expect(applied).toContain(eol + 'Work x' + eol);
                expect(helper.body.some(n => n.kind === 'VariableGroup')).toBe(false);
            }
        }
    });
    it('preserves multiple aliases and a read classified argument that a callee mutates', () => {
        const { applied } = extract('Sub Main(ByRef x As Long, ByRef y As Long)', ['x = x + 1', 'y = y + x', 'Mutate y'].join(eol), eol,
            ['Private Sub Mutate(ByRef value As Long)', 'value = 7', 'End Sub'].join(eol));
        expect(applied).toContain('Private Sub Work(ByRef x As Long, ByRef y As Long)');
        expect(applied).toContain(eol + 'Work x, y' + eol);
    });
    it.each([
        ['ByRef x() As Long', 'ReDim x(1 To 2)', 'ByRef x() As Long'],
        ['x() As Long', 'x(1) = 7', 'ByRef x() As Long'],
        ['Optional x As Variant', 'Debug.Print IsMissing(x)', 'ByRef x As Variant'],
        ['Optional x', 'Debug.Print IsMissing(x)', 'ByRef x'],
        ['Optional ByVal x As Long = 5', 'x = 7', 'ByRef x As Long'],
        ['Optional x As String = "as ="', 'x = "new"', 'ByRef x As String'],
        ['ByRef x%', 'x = 7', 'ByRef x%'],
        ['ByVal x$', 'x = "new"', 'ByRef x$'],
        ['ByRef [Case] As Long', '[Case] = 7', 'ByRef [Case] As Long'],
        ['ByRef x As VBA.Collection', 'Set x = New VBA.Collection', 'ByRef x As VBA.Collection'],
        ['ByRef x As [Collection]', 'Set x = New Collection', 'ByRef x As [Collection]'],
        ['ByVal x As Collection', 'Set x = New Collection', 'ByRef x As Collection'],
    ])('retains declaration semantics for %s', (parameter, body, binding) => {
        const { applied, helper } = extract('Sub Main(' + parameter + ')', body, eol);
        expect(applied).toContain('Private Sub Work(' + binding + ')');
        expect(helper.params).toHaveLength(1);
        expect(helper.params[0].optional).toBe(false);
        expect(helper.params[0].paramArray).toBe(false);
        expect(helper.params[0].byRef).toBe(true);
    });
    it.each(['ParamArray args() As Variant', 'ParamArray args()'])('refuses forwarding %s without emitting edits', parameter => {
        const prefix = ['Option Explicit', 'Sub Main(' + parameter + ')', ''].join(eol);
        const body = 'Debug.Print UBound(args)';
        const source = prefix + body + eol + 'End Sub' + eol;
        const result = extractMethod({ source, span: { start: prefix.length, end: prefix.length + body.length }, name: 'Work' });
        expect(result.ok).toBe(false);
        if (result.ok) throw Error('Expected refusal');
        expect(result.reason).toContain('ParamArray');
    });
    it('keeps an untyped parameter under the same module DefType directives', () => {
        const { applied } = extract('Sub Main(x)', 'x = 7', eol, 'DefLng A-Z');
        expect(applied).toContain('Private Sub Work(ByRef x)');
        expect(applied).toContain('DefLng A-Z');
    });
    it('reads continued type declarations without copying the continuation into the helper', () => {
        const { applied } = extract(['Sub Main(Optional ByVal x _', ' As _', ' Long = 5)'].join(eol), 'x = 7', eol);
        expect(applied).toContain('Private Sub Work(ByRef x As Long)');
    });
    it('keeps local Function outputs alongside parameter updates', () => {
        const { applied, helper } = extract('Sub Main(ByRef x As Long)', ['x = 7', 'total = x'].join(eol), eol, '', 'Debug.Print total' + eol, 'Dim total As Long');
        // The local belongs to Main, unlike its existing parameter binding.
        expect(helper.procKind).toBe('Function');
        expect(applied).toContain('Private Function Work(ByRef x As Long) As Long');
        expect(applied).toContain('total = Work(x)');
    });
});
