import { describe, expect, it } from 'vitest';
import { extractMethod } from '../src/analyzer/refactor/extractMethod';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';
import { analyzeModule } from '../src/analyzer/diagnostics/analyzeModule';
import { parseModule } from '../src/analyzer/parser/parseModule';
const errors=(source:string)=>analyzeModule(source,{knownIdentifiers:new Set<string>()}).filter(d=>d.code==='undeclared-variable');
function run(declaration:string,earlier:string,selected:string,eol:string){
    const prefix=['Option Explicit','Sub Main()',declaration,...(earlier?[earlier]:[]),''].join(eol);
    const source=prefix+selected+eol+'Debug.Print "after"'+eol+'End Sub'+eol;
    expect(errors(source)).toEqual([]);
    const result=extractMethod({source,span:{start:prefix.length,end:prefix.length+selected.length},name:'Work'});
    if(!result.ok)throw Error(result.reason);
    const text=applyVbaTextEdits(source,result.edits),at=text.indexOf('Private Sub Work');
    expect(errors(text)).toEqual([]); expect(text.slice(0,at)).toContain('Work');
    expect(text.slice(0,at)).toContain('Debug.Print "after"');
    if(earlier)expect(text.slice(0,at)).toContain(earlier);
    const decls=(name:string)=>{
        const procedure=parseModule(text).members.find(n=>n.kind==='Procedure'&&n.name===name);
        if(procedure?.kind!=='Procedure')throw Error('missing '+name);
        return procedure.body.filter(n=>n.kind==='VariableGroup').flatMap(n=>n.declarations.map(d=>d.name));
    };
    return {text,caller:text.slice(0,at),helper:text.slice(at),callerNames:decls('Main'),helperNames:decls('Work')};
}
describe.each(['\n','\r\n','\r'])('earlier caller uses with %j',eol=>{
    it.each(['Debug.Print total','total=2','If True Then'+eol+'Debug.Print total'+eol+'End If'])('keeps a declaration needed by %s',earlier=>{
        const out=run('Dim total As Long',earlier,'total=3',eol);
        expect(out.callerNames).toEqual(['total']); expect(out.helperNames).toEqual(['total']);
    });
    it.each(['a','b','c'])('keeps only the caller-used clause %s from a moved group',name=>{
        const out=run('Dim a As Long, b As Long, c As Long','Debug.Print '+name,['a=1','b=2','c=3'].join(eol),eol);
        expect(out.callerNames).toEqual([name]); expect([...out.helperNames].sort()).toEqual(['a','b','c']);
    });
    it('keeps multiple caller-used clauses and an untouched sibling',()=>{
        const out=run('Dim a As Long, b As Long, c As Long, unused As String','Debug.Print a, c',['a=1','b=2','c=3'].join(eol),eol);
        expect(out.callerNames).toEqual(['a','c','unused']); expect([...out.helperNames].sort()).toEqual(['a','b','c']);
    });
    it('preserves declaration type suffix and fixed-length syntax in both scopes',()=>{
        const out=run('Dim count%, shortText As String * 5','Debug.Print count, shortText','count=3'+eol+'shortText="hello"',eol);
        expect(out.caller).toContain('Dim count%, shortText As String * 5');
        expect(out.helper).toContain('Dim count%'); expect(out.helper).toContain('Dim shortText As String * 5');
    });
    it('still removes a caller declaration with no earlier use',()=>{
        const out=run('Dim total As Long','','total=3',eol);
        expect(out.callerNames).toEqual([]); expect(out.helperNames).toEqual(['total']);
    });
    it('does not treat comments or string literals as earlier uses',()=>{
        const out=run('Dim total As Long',"' total"+eol+'Debug.Print "total"','total=3',eol);
        expect(out.callerNames).toEqual([]); expect(out.helperNames).toEqual(['total']);
    });
});
