import { describe, expect, it } from 'vitest';
import { extractMethod } from '../src/analyzer/refactor/extractMethod';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';
import { analyzeModule } from '../src/analyzer/diagnostics/analyzeModule';
import { parseModule } from '../src/analyzer/parser/parseModule';
const cases = [
    ['Dim total As Long', 'total', '3', ''],
    ['Dim message As String', 'message', '"hello"', ''],
    ['Dim ratio As Double', 'ratio', '1.5', ''],
    ['Dim flag As Boolean', 'flag', 'True', ''],
    ['Dim untyped', 'untyped', '3', ''],
    ['Dim shortText As String * 5', 'shortText', '"hello"', ''],
    ['Dim count%', 'count', '3', ''],
    ['', 'number', '3', 'ByRef number As Long'],
    ['', 'number', '3', 'ByVal number%'],
] as const;
const undeclared = (source: string) => analyzeModule(source, {knownIdentifiers: new Set<string>()}).filter(d=>d.code==='undeclared-variable');
describe.each(['\n','\r\n','\r'])('Function output declaration with %j', eol => {
    it.each(cases)('declares the helper output from %s (%s)', (declaration, name, value, parameter) => {
        const prefix=['Option Explicit','Sub Main('+parameter+')',...(declaration?[declaration]:[]),''].join(eol);
        const selection=name+' = '+value;
        const source=prefix+selection+eol+'Debug.Print '+name+eol+'End Sub'+eol;
        expect(undeclared(source)).toEqual([]);
        const result=extractMethod({source,span:{start:prefix.length,end:prefix.length+selection.length},name:'Work'});
        if(!result.ok)throw Error(result.reason);
        const applied=applyVbaTextEdits(source,result.edits);
        expect(applied).toContain('Private Function Work()');
        expect(undeclared(applied)).toEqual([]);
        const helper=parseModule(applied).members.find(n=>n.kind==='Procedure'&&n.name==='Work');
        expect(helper?.kind).toBe('Procedure');
        if(helper?.kind!=='Procedure')throw Error('missing helper');
        const declared=helper.body.filter(n=>n.kind==='VariableGroup').flatMap(n=>n.declarations);
        expect(declared.map(d=>d.name)).toEqual([name]);
        if(declaration)expect(applied.slice(applied.indexOf('Private Function Work'))).toContain(declaration);
        expect(applied.slice(0,applied.indexOf('Private Function Work'))).toContain(name+' = Work()');
    });
    it('uses the implicit result variable when the helper has the output name', () => {
        const prefix=['Option Explicit','Sub Main()','Dim total As Long',''].join(eol), selection='total=1';
        const source=prefix+selection+eol+'Debug.Print total'+eol+'End Sub'+eol;
        const result=extractMethod({source,span:{start:prefix.length,end:prefix.length+selection.length},name:'TOTAL'});
        if(!result.ok)throw Error(result.reason);
        const applied=applyVbaTextEdits(source,result.edits);
        expect(undeclared(applied)).toEqual([]);
        expect(applied.slice(applied.indexOf('Private Function TOTAL'))).not.toContain('Dim total');
    });
    it('does not redeclare ByRef outputs in an extracted Sub', () => {
        const prefix=['Option Explicit','Sub Main()','Dim a As Long, b As Long',''].join(eol), selection='a=1'+eol+'b=2';
        const source=prefix+selection+eol+'Debug.Print a, b'+eol+'End Sub'+eol;
        const result=extractMethod({source,span:{start:prefix.length,end:prefix.length+selection.length},name:'Work'});
        if(!result.ok)throw Error(result.reason);
        const applied=applyVbaTextEdits(source,result.edits);
        expect(applied).toContain('Private Sub Work(ByRef'); expect(undeclared(applied)).toEqual([]);
        expect(applied.slice(applied.indexOf('Private Sub Work'))).not.toContain('Dim ');
    });
});
