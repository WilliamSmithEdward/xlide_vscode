import { describe, expect, it } from 'vitest';
import { introduceParameter } from '../src/analyzer/refactor/introduceParameter';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';
for(const eol of ['\n','\r\n','\r']) {
 describe('recursive receiver binding '+JSON.stringify(eol),()=>{
  for(const [body,expected] of [
   [['If depth > 0 Then Me.Report depth - 1'], ['If depth > 0 Then Me.Report depth - 1, 3']],
   [['With Me', '.Report depth - 1', 'End With'], ['With Me', '.Report depth - 1, 3', 'End With']],
   [['Dim target As Module1', 'target.Report depth - 1'], ['Dim target As Module1', 'target.Report depth - 1, 3']],
   [['Dim other As Other', 'With other', '.Report depth - 1', 'End With', 'Report depth - 1'], ['Dim other As Other', 'With other', '.Report depth - 1', 'End With', 'Report depth - 1, 3']],
   [['Dim wrapper As Holder', 'wrapper.Child.Report depth - 1'], ['Dim wrapper As Holder', 'wrapper.Child.Report depth - 1, 3']],
   [['Dim wrapper As Holder', 'With wrapper.Child', '.Report depth - 1', 'End With'], ['Dim wrapper As Holder', 'With wrapper.Child', '.Report depth - 1, 3', 'End With']],
   [['Dim other As Other', 'With Me', 'With other', '.Report depth - 1', 'End With', '.Report depth - 1', 'End With'], ['Dim other As Other', 'With Me', 'With other', '.Report depth - 1', 'End With', '.Report depth - 1, 3', 'End With']],
   [['Me.[Report] depth - 1'], ['Me.[Report] depth - 1, 3']],
   [['With Me', '.[Report] depth - 1', 'End With'], ['With Me', '.[Report] depth - 1, 3', 'End With']],
  ]) {
   it(body.join(' / '),()=>{
    const source=['Public Sub Report(ByVal depth As Long)','Dim limit As Long','limit = 3',...body,'End Sub',''].join(eol);
    const result=introduceParameter({source,offset:source.indexOf('limit'),moduleName:'Module1',moduleKinds:{Module1:'class',Other:'class',Holder:'class'},otherModuleSources:{Other:['Public Sub Report(ByVal depth As Long)','Debug.Print depth','End Sub',''].join(eol),Holder:'Public Child As Module1'+eol}});
    if(!result.ok)throw new Error(result.reason);
    expect(applyVbaTextEdits(source,result.edits)).toBe(['Public Sub Report(ByVal depth As Long, ByVal limit As Long)',...expected,'End Sub',''].join(eol));
    expect(result.otherModules).toBeUndefined();
   });
  }
  for(const body of [['Me.Report depth - 1'],['With Me','.Report depth - 1','End With']]) {
   it('updates private recursion '+body.join(' / '),()=>{
    const source=['Private Sub Report(ByVal depth As Long)','Dim limit As Long','limit = 3',...body,'End Sub',''].join(eol);
    const result=introduceParameter({source,offset:source.indexOf('limit'),moduleName:'Module1',moduleKinds:{Module1:'class'}});
    if(!result.ok)throw new Error(result.reason);
    expect(applyVbaTextEdits(source,result.edits)).toBe(['Private Sub Report(ByVal depth As Long, ByVal limit As Long)',...body.map(line=>line.includes('.Report')?line+', 3':line),'End Sub',''].join(eol));
   });
  }
  it('refuses unresolved late-bound receivers instead of leaving potential recursion stale',()=>{
   const source=['Public Sub Report(ByVal receiver As Object)','Dim limit As Long','limit = 3','receiver.Report receiver','End Sub',''].join(eol);
   const result=introduceParameter({source,offset:source.indexOf('limit'),moduleName:'Module1',moduleKinds:{Module1:'class'}});
   expect(result.ok).toBe(false);
  });
  it('rejects an initializer that would recursively invoke the old signature',()=>{
   const source=['Public Function Report(ByVal depth As Long) As Long','If depth = 0 Then Report = 2: Exit Function','Dim limit As Long','limit = Report(0)','Report = limit','End Function',''].join(eol);
   const result=introduceParameter({source,offset:source.indexOf('limit'),moduleName:'Module1'});
   expect(result.ok).toBe(false);
  });
  it('keeps module-name shadowing ahead of module qualification',()=>{
   const source=['Public Sub Report(ByVal Module1 As Other)','Dim limit As Long','limit = 3','Module1.Report 1','Report Module1','End Sub',''].join(eol);
   const result=introduceParameter({source,offset:source.indexOf('limit'),moduleName:'Module1',moduleKinds:{Other:'class'},otherModuleSources:{Other:'Public Sub Report(ByVal depth As Long)'+eol+'End Sub'+eol}});
   if(!result.ok)throw new Error(result.reason);
   expect(applyVbaTextEdits(source,result.edits)).toBe(['Public Sub Report(ByVal Module1 As Other, ByVal limit As Long)','Module1.Report 1','Report Module1, 3','End Sub',''].join(eol));
  });
 });
}
