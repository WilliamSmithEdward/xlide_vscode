import { describe, expect, it } from 'vitest';
import { callSitesOf } from '../src/analyzer/refactor/callSites';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';
for(const eol of ['\n','\r\n','\r']) {
 describe('call sites preserve non-call syntax '+JSON.stringify(eol),()=>{
  for(const statement of ['Report:'+eol+'GoTo Report','Other Report:=1','x = AddressOf Report','Debug.Print value!Report','Set value = New Report','Dim value As Report','If TypeOf value Is Report Then Debug.Print 1']) {
   it(statement,()=>expect(callSitesOf(['Sub Caller()',statement,'End Sub',''].join(eol),'Report')).toEqual([]));
  }
  for(const [before,after] of [
   ['If True Then Report: Report 1','If True Then Report 3: Report 1, 3'],
   ['If TypeOf value Is Report And value Is Report Then Debug.Print 1','If TypeOf value Is Report And value Is Report(3) Then Debug.Print 1'],
   ['Call Report: Report 1','Call Report(3): Report 1, 3'],
   ['Module1.Report: Report 1','Module1.Report 3: Report 1, 3'],
  ]) {
   it(before,()=>{
    const source=['Sub Caller()',before,'End Sub',''].join(eol);
    const sites=callSitesOf(source,'Report');
    expect(sites).toHaveLength(before.startsWith('If TypeOf')?1:2);
    expect(applyVbaTextEdits(source,sites.map(s=>({span:s.argumentInsert,newText:s.argumentText('3')})))).toBe(source.replace(before,after));
   });
  }
 });
}
