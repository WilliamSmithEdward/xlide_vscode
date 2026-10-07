import { describe, expect, it } from 'vitest';
import { extractMethod } from '../src/analyzer/refactor/extractMethod';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { walkBody } from '../src/analyzer/refactor/shared';
function extract(declaration: string, selection: string, tail: string, eol: string) {
    const prefix = ['Option Explicit', 'Sub Main()', declaration, ''].join(eol);
    const source = prefix + selection + eol + tail + eol + 'End Sub' + eol;
    const result = extractMethod({ source, span: { start: prefix.length, end: prefix.length + selection.length }, name: 'Work' });
    if (!result.ok) throw Error(result.reason);
    const edits = [...result.edits].sort((a,b) => a.span.start - b.span.start);
    for (let i=1;i<edits.length;i++) expect(edits[i].span.start, 'edits must not overlap').toBeGreaterThanOrEqual(edits[i-1].span.end);
    const applied = applyVbaTextEdits(source, result.edits);
    const caller = applied.slice(0, applied.indexOf('Private Sub Work'));
    expect(caller).toContain('Work'); expect(caller).toContain(tail);
    return { applied, caller, helper: applied.slice(applied.indexOf('Private Sub Work')), result };
}
describe.each(['\n','\r\n','\r'])('grouped declaration extraction with %j', eol => {
    it('removes a shared declaration once and preserves the call and later code', () => {
        const {caller,helper,result}=extract('Dim a As Long, b As Long', 'a=1'+eol+'b=2', 'Debug.Print "after"', eol);
        expect(caller).not.toContain('Dim'); expect(helper).toContain('Dim a As Long'); expect(helper).toContain('Dim b As Long');
        expect(result.edits.filter(edit=>edit.newText==='')).toHaveLength(1);
    });
    it.each(['a','b','c'])('moves only %s and preserves untouched sibling declarations', name => {
        const {caller,helper}=extract('Dim a As Long, b As String * 5, c%', name+' = '+(name==='b'?'"x"':'1'), 'Debug.Print "after"', eol);
        const module=parseModule(caller);
        const procedure=module.members.find(node=>node.kind==='Procedure')!;
        if(procedure.kind!=='Procedure')throw Error('missing caller');
        const names=[...walkBody(procedure.body)].filter(node=>node.kind==='VariableGroup').flatMap(node=>node.kind==='VariableGroup'?node.declarations.map(d=>d.name):[]);
        expect(names).toEqual(['a','b','c'].filter(n=>n!==name));
        expect(helper).toContain(name==='a'?'Dim a As Long':name==='b'?'Dim b As String * 5':'Dim c%');
        if(name!=='b')expect(caller).toContain('b As String * 5');
        if(name!=='c')expect(caller).toContain('c%');
    });
    it('preserves unrelated statements on a declaration line', () => {
        const {caller}=extract('Dim a As Long : Dim b As Long : Debug.Print "before"', 'a=1'+eol+'b=2', 'Debug.Print "after"', eol);
        expect(caller).toContain('Debug.Print "before"'); expect(caller).not.toContain('Dim');
    });
    it('unions overlapping separator removals for adjacent fully moved statements', () => {
        const {caller}=extract('Dim a As Long : Dim b As Long', 'a=1'+eol+'b=2', 'Debug.Print "after"', eol);
        expect(caller).not.toContain('Dim'); expect(caller).not.toContain(':');
    });
    it('retains a selected input declaration in the caller and uses only its parameter in the helper', () => {
        const {caller,helper}=extract('', 'Dim a As Long'+eol+'Debug.Print a', 'Debug.Print "after"', eol);
        expect(caller).toContain('Dim a As Long'); expect(helper).toContain('ByRef a As Long');
        expect(helper).not.toContain('Dim a');
    });
    it('preserves a sibling declaration in a separate colon statement', () => {
        const {caller}=extract('Debug.Print "before" : Dim a As Long : Dim b As Long', 'a=1', 'Debug.Print b', eol);
        expect(caller).toContain('Debug.Print "before"'); expect(caller).toContain('Dim b As Long');
    });
    it('retains an array binding and its object sibling in the caller', () => {
        const {caller,helper}=extract('Dim arr(0 To 2) As Long, item As Collection', 'arr(0)=1'+eol+'Set item = New Collection', 'Debug.Print "after"', eol);
        expect(caller).toContain('Dim arr(0 To 2) As Long');
        expect(helper).toContain('ByRef arr() As Long'); expect(helper).not.toContain('Dim arr');
        expect(caller).toContain('item As Collection');
        expect(helper).toContain('ByRef item As Collection'); expect(helper).not.toContain('Dim item');
    });
    it('does not duplicate a declaration already inside the selected block', () => {
        const {caller,helper}=extract('', 'Dim a As Long, b As Long'+eol+'a=1'+eol+'b=2', 'Debug.Print "after"', eol);
        expect(caller).not.toContain('Dim'); expect(helper.match(/Dim /g)).toHaveLength(1);
    });
    it('retains caller siblings when their shared declaration is selected too', () => {
        const {caller,helper}=extract('', 'Dim a As Long, b As Long'+eol+'a=1', 'Debug.Print b', eol);
        expect(caller).toContain('Dim b As Long'); expect(caller).not.toContain('a As Long');
        expect(helper).toContain('Dim a As Long'); expect(helper).not.toContain('b As Long');
    });
    it('keeps a selected output declaration in both caller and extracted Function', () => {
        const prefix=['Option Explicit','Sub Main()',''].join(eol), selection='Dim a As Long, b As Long'+eol+'a=1'+eol+'b=2';
        const source=prefix+selection+eol+'Debug.Print a'+eol+'End Sub'+eol;
        const result=extractMethod({source,span:{start:prefix.length,end:prefix.length+selection.length},name:'Work'});
        if(!result.ok)throw Error(result.reason);
        const text=applyVbaTextEdits(source,result.edits), at=text.indexOf('Private Function Work');
        expect(text.slice(0,at)).toContain('Dim a As Long'); expect(text.slice(0,at)).not.toContain('b As Long');
        expect(text.slice(at)).toContain('Dim a As Long, b As Long');
        expect(text.slice(0,at)).toContain('a = Work()'); expect(text.slice(0,at)).toContain('Debug.Print a');
    });
    it('keeps declaration comments and continuation syntax for remaining siblings', () => {
        const declaration='Dim a As Long, _'+eol+'    b As String * 5 '+"' keep this comment";
        const {caller}=extract(declaration, 'a=1', 'Debug.Print b', eol);
        expect(caller).toContain('Dim b As String * 5'); expect(caller).toContain("' keep this comment");
    });
});
