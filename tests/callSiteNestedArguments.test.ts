import { describe, expect, it } from 'vitest';
import { callSitesOf } from '../src/analyzer/refactor/callSites';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';
import { introduceParameter } from '../src/analyzer/refactor/introduceParameter';
for (const eol of ['\n','\r\n','\r']) {
 describe('nested and continued call fidelity '+JSON.stringify(eol), () => {
  for (const [before,after] of [
   ['Go Go(1)', 'Go Go(1, 3), 3'],
   ['x = Go(Go(1))', 'x = Go(Go(1, 3), 3)'],
   ['Call Go', 'Call Go(3)'],
   ['x = Go + 1', 'x = Go(3) + 1'],
   ['If Go = 1 Then Debug.Print 2', 'If Go(3) = 1 Then Debug.Print 2'],
   ['[Go] "hello"', '[Go] "hello", 3'],
   ['Go first:=1', 'Go first:=1, added:=3'],
   ['Call Go(first:=1)', 'Call Go(first:=1, added:=3)'],
   ['Go Other(first:=1)', 'Go Other(first:=1), 3'],
   ['[Module1].Go (1)', '[Module1].Go (1), 3'],
   ['If True Then Go 1: Go 2 Else Go 4', 'If True Then Go 1, 3: Go 2, 3 Else Go 4, 3'],
  ]) {
   it(before, () => {
    const source = ['Option Explicit', 'Sub Caller()', before, 'End Sub',''].join(eol);
    const sites = callSitesOf(source,'Go');
    expect(sites.length).toBeGreaterThan(0);
    expect(applyVbaTextEdits(source,sites.map(site=>({span:site.argumentInsert,newText:site.argumentText('3','added')}))))
     .toBe(source.replace(before,after));
   });
  }
  for (const [before,after] of [
   [['Go _',' "hello": Debug.Print 2'], ['Go _',' "hello", 3: Debug.Print 2']],
   [['Call Go( _',' 1)'], ['Call Go( _',' 1, 3)']],
   [['x = _',' Go(1)'], ['x = _',' Go(1, 3)']],
  ]) {
   it('preserves continuation '+before[0], () => {
    const source = ['Option Explicit','Sub Caller()',...before,'End Sub',''].join(eol);
    const sites=callSitesOf(source,'Go');
    expect(sites).toHaveLength(1);
    expect(applyVbaTextEdits(source,sites.map(site=>({span:site.argumentInsert,newText:site.argumentText('3')}))))
     .toBe(source.replace(before.join(eol),after.join(eol)));
   });
  }
  it('leaves assignment and loop-control targets alone',()=>{
   for(const statement of ['Go = 1','Let Go = 1','Set Go = Nothing','For Go = 1 To 2']){
    const source=['Sub Caller()',statement,'End Sub'].join(eol);
    expect(callSitesOf(source,'Go')).toEqual([]);
   }
  });
  it('adds the declared name at actual named argument calls', () => {
   const source=['Public Sub Report(ByVal title As String)','Dim limit As Long','limit = 3','Debug.Print limit','End Sub','Sub Caller()', 'Report title:="hello"', 'Call Report(title:="other")', 'End Sub',''].join(eol);
   const result=introduceParameter({source,offset:source.indexOf('limit'),moduleName:'Module1'});
   if(!result.ok)throw new Error(result.reason);
   const output=applyVbaTextEdits(source,result.edits);
   expect(output).toContain('Report title:="hello", limit:=3');
   expect(output).toContain('Call Report(title:="other", limit:=3)');
  });
 });
}
