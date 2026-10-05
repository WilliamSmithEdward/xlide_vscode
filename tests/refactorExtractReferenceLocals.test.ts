import { describe, expect, it } from 'vitest';
import { extractMethod } from '../src/analyzer/refactor/extractMethod';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';
import { analyzeModule } from '../src/analyzer/diagnostics/analyzeModule';
import { parseModule } from '../src/analyzer/parser/parseModule';
function fixture(decl: string, body: string, eol: string, before = '', after = '', module = '', name = 'Work') {
    const convert = (s: string) => s.replace(/\n/g, eol);
    const prefix = convert('Option Explicit\n' + (module ? module + '\n' : '') + 'Sub Main()\n' + (decl ? decl + '\n' : '') + (before ? before + '\n' : ''));
    const selected = convert(body);
    return { source: prefix + selected + eol + (after ? convert(after) + eol : '') + 'End Sub' + eol, span: { start: prefix.length, end: prefix.length + selected.length }, name };
}
function applied(input: ReturnType<typeof fixture>) {
    const result = extractMethod(input);
    if (!result.ok) throw Error(result.reason);
    const source = applyVbaTextEdits(input.source, result.edits);
    const helper = parseModule(source).members.find(n => n.kind === 'Procedure' && n.name === input.name);
    if (helper?.kind !== 'Procedure') throw Error('Missing helper');
    expect(analyzeModule(source, { knownIdentifiers: new Set<string>() }).filter(d => d.code === 'undeclared-variable')).toEqual([]);
    return { source, helper, caller: source.slice(0, source.indexOf('Private Sub Work')) };
}
describe.each(['\n', '\r\n', '\r'])('Extract Method reference-capable locals with %j', eol => {
    it.each([
        ['Dim x As Collection', 'Set x = New Collection', 'ByRef x As Collection', ''],
        ['Dim x As Object', 'Set x = New Collection', 'ByRef x As Object', ''],
        ['Dim x As Variant', 'Set x = New Collection', 'ByRef x As Variant', ''],
        ['Dim x As Variant', 'x = 7', 'ByRef x As Variant', ''],
        ['Dim x', 'x = 7', 'ByRef x', ''],
        ['Dim x As Variant', 'x = Array(3, 7)', 'ByRef x As Variant', ''],
        ['Dim x As Variant', 'Debug.Print x', 'ByRef x As Variant', ''],
        ['Dim x As Collection', 'x.Add 7', 'ByRef x As Collection', ''],
        ['Dim x As VBA.Collection', 'Set x = New VBA.Collection', 'ByRef x As VBA.Collection', ''],
        ['Dim x As [Collection]', 'Set x = New Collection', 'ByRef x As [Collection]', ''],
        ['Dim x As Shade', 'x = Blue', 'ByRef x As Shade', 'Private Enum Shade\nBlue = 7\nEnd Enum'],
        ['Dim x As Pair', 'x.Count = 7', 'ByRef x As Pair', 'Private Type Pair\nCount As Long\nEnd Type'],
    ])('retains %s when extracting %s', (decl, body, binding, module) => {
        for (const after of ['', body.includes('x.Count') ? 'Debug.Print x.Count' : 'Debug.Print x']) {
            const { source, helper, caller } = applied(fixture(decl, body, eol, '', after, module));
            expect(source).toContain('Private Sub Work(' + binding + ')');
            expect(helper.params).toHaveLength(1); expect(helper.params[0].byRef).toBe(true);
            expect(helper.body.some(n => n.kind === 'VariableGroup')).toBe(false);
            expect(caller).toContain(decl); expect(caller).toContain('Work x');
            expect(source.slice(source.indexOf('Private Sub Work'))).not.toContain('Work = x');
        }
    });
    it.each(['Variant', 'Collection'])('keeps indirect ByRef writes to a %s input', type => {
        const assignment = type === 'Variant' ? 'p = 7' : 'Set p = New Collection';
        const { source } = applied(fixture('Dim x As ' + type, 'Mutate x', eol, '', 'Debug.Print x', 'Private Sub Mutate(ByRef p As ' + type + ')\n' + assignment + '\nEnd Sub'));
        expect(source).toContain('Private Sub Work(ByRef x As ' + type + ')');
        expect(source).toContain('Mutate x');
    });
    it('retains an untyped local under its original DefType', () => {
        const { source } = applied(fixture('Dim x', 'x = 7', eol, '', 'Debug.Print x', 'DefLng X-X'));
        expect(source).toContain('Private Sub Work(ByRef x)');
    });
    it('keeps selected reference declarations in the caller', () => {
        const { source, caller } = applied(fixture('', 'Dim x As Collection\nSet x = New Collection', eol, '', 'Debug.Print x.Count'));
        expect(caller).toContain('Dim x As Collection'); expect(caller).toContain('Work x');
        expect(source.slice(source.indexOf('Private Sub Work'))).not.toContain('Dim x');
    });
    it('retains a primitive Function output alongside an object binding', () => {
        const { source, helper } = applied(fixture('Dim x As Collection, n As Long', 'Set x = New Collection\nx.Add 7\nn = 8', eol, '', 'Debug.Print x.Count, n'));
        expect(helper.procKind).toBe('Function');
        expect(source).toContain('Private Function Work(ByRef x As Collection) As Long');
        expect(source).toContain('n = Work(x)');
    });
    it.each(['Dim x As New Collection', 'Dim x() As New Collection'])('refuses automatic creation that a formal cannot retain: %s', decl => {
        const result = extractMethod(fixture(decl, 'Set x = Nothing', eol));
        expect(result.ok).toBe(false); if (result.ok) throw Error('Expected refusal');
        expect(result.reason).toContain('As New');
    });
    it('allows an unrelated primitive selection with an As New declaration', () => {
        const { source } = applied(fixture('Dim x As New Collection\nDim n As Long', 'Debug.Print n', eol));
        expect(source).toContain('Private Sub Work(ByVal n As Long)');
    });
    it.each(['Dim x As Variant', 'Dim x As Collection'])('refuses a helper/parameter collision for %s', decl => {
        const result = extractMethod(fixture(decl, 'Debug.Print x', eol, '', '', '', 'X'));
        expect(result.ok).toBe(false); if (result.ok) throw Error('Expected refusal');
        expect(result.reason).toContain('parameter binding');
    });
    it.each(['obj.x.Add 7', 'obj!x.Add 7', 'obj.[x].Add 7', 'With obj\n.x.Add 7\nEnd With', 'With obj\n!x.Add 7\nEnd With'])('does not bind a same-named member in %s', body => {
        const { source } = applied(fixture('Dim x As New Collection\nDim obj As Collection', body, eol));
        expect(source).toContain('Private Sub Work(ByRef obj As Collection)');
        expect(source).not.toContain('ByRef x');
    });

});
