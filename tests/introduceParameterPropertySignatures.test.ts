import { describe, expect, it } from 'vitest';
import { introduceParameter } from '../src/analyzer/refactor/introduceParameter';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';
for(const eol of ['\n','\r\n','\r']) {
 describe('Introduce Parameter property signatures '+JSON.stringify(eol),()=>{
  it('preserves valid unpaired getter reads and updates callers',()=>{
   const source=['Public Property Get Report() As Long','Dim limit As Long','limit = 3','Report = limit','Debug.Print Report','End Property','Sub Caller()','Debug.Print Report','End Sub',''].join(eol);
   const result=introduceParameter({source,offset:source.indexOf('limit'),moduleName:'Module1'});
   if(!result.ok)throw new Error(result.reason);
   expect(applyVbaTextEdits(source,result.edits)).toBe(['Public Property Get Report(ByVal limit As Long) As Long','Report = limit','Debug.Print Report','End Property','Sub Caller()','Debug.Print Report(3)','End Sub',''].join(eol));
  });
  for(const accessor of ['Let','Set']) {
   it('does not break a Get/'+accessor+' accessor family',()=>{
    const type=accessor==='Let'?'Long':'Object';
    const source=['Public Property Get Report() As '+type,'Dim limit As Long','limit = 3',accessor==='Let'?'Report = limit':'Set Report = Nothing','End Property','Public Property '+accessor+' Report(ByVal value As '+type+')','End Property',''].join(eol);
    expect(introduceParameter({source,offset:source.indexOf('limit'),moduleName:'Module1'}).ok).toBe(false);
   });
   it('does not shift a '+accessor+' setter value argument',()=>{
    const type=accessor==='Let'?'Long':'Object';
    const source=['Public Property '+accessor+' Report(ByVal value As '+type+')','Dim limit As Long','limit = 3','Debug.Print limit','End Property',''].join(eol);
    expect(introduceParameter({source,offset:source.indexOf('limit'),moduleName:'Module1'}).ok).toBe(false);
   });
  }
 });
}
