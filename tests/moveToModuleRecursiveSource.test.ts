import { describe, expect, it } from 'vitest';
import { moveToModule } from '../src/analyzer/refactor/moveToModule';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';
for(const eol of ['\n','\r\n','\r']) describe('moved recursion '+JSON.stringify(eol),()=>{
 for(const call of ['If depth > 0 Then Reports.Build depth - 1','If depth > 0 Then Call Reports.Build(depth - 1)','If depth > 0 Then Reports.Build depth - 1: Reports.Build 0','If depth > 0 Then rEpOrTs . Build depth - 1']) {
  it(call,()=>{
   const comment="''' Reports.Build recursion";
   const moved=[comment,'Public Sub Build(ByVal depth As Long)',call,'Debug.Print "Reports.Build"',"' Reports.Build",'End Sub'].join(eol);
   const source=['Option Explicit','',moved,'','Public Sub Retained()', 'Reports.Build 1','End Sub',''].join(eol);
   const target=['Option Explicit','Public Sub Existing()','Reports.Build 2','End Sub',''].join(eol);
   const result=moveToModule({source,offset:source.indexOf('Public Sub Build'),moduleName:'Reports',targetModuleName:'LongerHelpers',otherModuleSources:{LongerHelpers:target}});
   if(!result.ok)throw new Error(result.reason);
   for(const edits of [result.edits,...(result.otherModules??[]).map(module=>module.edits)]) {
    const ordered=[...edits].sort((a,b)=>a.span.start-b.span.start);
    for(let i=1;i<ordered.length;i++)expect(ordered[i-1].span.end).toBeLessThanOrEqual(ordered[i].span.start);
   }
   expect(applyVbaTextEdits(source,result.edits)).toBe(['Option Explicit','','Public Sub Retained()','LongerHelpers.Build 1','End Sub',''].join(eol));
   const changedCall=call.replace(/Reports|rEpOrTs/g,'LongerHelpers');
   const expectedMoved=[comment,'Public Sub Build(ByVal depth As Long)',changedCall,'Debug.Print "Reports.Build"',"' Reports.Build",'End Sub'].join(eol);
   expect(result.otherModules).toHaveLength(1);
   expect(applyVbaTextEdits(target,result.otherModules![0].edits)).toBe(['Option Explicit','Public Sub Existing()','LongerHelpers.Build 2','End Sub','',expectedMoved,''].join(eol));
  });
 }
});
