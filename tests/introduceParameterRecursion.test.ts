import { describe, expect, it } from 'vitest';
import { introduceParameter } from '../src/analyzer/refactor/introduceParameter';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';

for (const eol of ['\n', '\r\n', '\r']) {
 describe('Introduce Parameter recursion '+JSON.stringify(eol), () => {
  for (const [before,after] of [
   [['10 Report depth - 1'], ['10 Report depth - 1, 3']],
   [['10 Call Report(depth - 1)'], ['10 Call Report(depth - 1, 3)']],
   [['If depth > 0 Then Report depth - 1'], ['If depth > 0 Then Report depth - 1, 3']],
   [['If depth > 0 Then Call Report(depth - 1)'], ['If depth > 0 Then Call Report(depth - 1, 3)']],
   [['If depth > 0 Then Module1.Report depth - 1'], ['If depth > 0 Then Module1.Report depth - 1, 3']],
   [['If depth > 0 Then Report depth:=depth - 1'], ['If depth > 0 Then Report depth:=depth - 1, limit:=3']],
   [['If depth > 0 Then Report _', ' depth - 1'], ['If depth > 0 Then Report _', ' depth - 1, 3']],
   [['If depth > 0 Then Report depth - 1: Report depth - 1'], ['If depth > 0 Then Report depth - 1, 3: Report depth - 1, 3']],
  ]) {
   it('updates recursion '+before[0], () => {
    const source=['Public Sub Report(ByVal depth As Long)', 'Dim limit As Long', 'limit = 3', 'Debug.Print limit', ...before, 'End Sub', ''].join(eol);
    const result=introduceParameter({source,offset:source.indexOf('limit'),moduleName:'Module1'});
    if(!result.ok)throw new Error(result.reason);
    expect(applyVbaTextEdits(source,result.edits)).toBe(['Public Sub Report(ByVal depth As Long, ByVal limit As Long)', 'Debug.Print limit', ...after, 'End Sub', ''].join(eol));
   });
  }
  for (const [before,after] of [
   ['If depth > 0 Then Report = Report(depth - 1)', 'If depth > 0 Then Report = Report(depth - 1, 3)'],
   ['If depth > 0 Then Report = Module1.Report(depth - 1)', 'If depth > 0 Then Report = Module1.Report(depth - 1, 3)'],
   ['If depth > 0 Then Call Report(depth - 1)', 'If depth > 0 Then Call Report(depth - 1, 3)'],
  ]) {
   it('preserves result variables while updating '+before, () => {
    const source=['Public Function Report(ByVal depth As Long) As Long', 'Dim limit As Long', 'limit = 3', 'Report = limit', before, 'Debug.Print Report', 'Report = Report + 1', 'End Function', ''].join(eol);
    const result=introduceParameter({source,offset:source.indexOf('limit'),moduleName:'Module1'});
    if(!result.ok)throw new Error(result.reason);
    expect(applyVbaTextEdits(source,result.edits)).toBe(['Public Function Report(ByVal depth As Long, ByVal limit As Long) As Long', 'Report = limit', after, 'Debug.Print Report', 'Report = Report + 1', 'End Function', ''].join(eol));
   });
  }
  it('retains nonrecursive function result reads', () => {
   const source=['Public Function Report() As Long', 'Dim limit As Long', 'limit = 3', 'Report = limit', 'Debug.Print Report', 'Report = Report + 1', 'End Function', ''].join(eol);
   const result=introduceParameter({source,offset:source.indexOf('limit'),moduleName:'Module1'});
   if(!result.ok)throw new Error(result.reason);
   expect(applyVbaTextEdits(source,result.edits)).toBe(['Public Function Report(ByVal limit As Long) As Long', 'Report = limit', 'Debug.Print Report', 'Report = Report + 1', 'End Function', ''].join(eol));
  });
  it('does not update another qualified owner in the recursive body', () => {
   const source=['Public Sub Report(ByVal depth As Long)', 'Dim limit As Long', 'limit = 3', 'If depth > 0 Then Report depth - 1', 'Module2.Report depth', 'End Sub', ''].join(eol);
   const result=introduceParameter({source,offset:source.indexOf('limit'),moduleName:'Module1',otherModuleSources:{Module2:['Public Sub Report(ByVal depth As Long)', 'Debug.Print depth', 'End Sub', ''].join(eol)}});
   if(!result.ok)throw new Error(result.reason);
   expect(applyVbaTextEdits(source,result.edits)).toBe(['Public Sub Report(ByVal depth As Long, ByVal limit As Long)', 'If depth > 0 Then Report depth - 1, 3', 'Module2.Report depth', 'End Sub', ''].join(eol));
   expect(result.otherModules).toBeUndefined();
  });
  it('refuses moving an initializer that reads the function result variable', () => {
   const source=['Public Function Report() As Long', 'Report = 2', 'Dim limit As Long', 'limit = Report', 'Report = limit + 1', 'End Function', ''].join(eol);
   const result=introduceParameter({source,offset:source.indexOf('limit'),moduleName:'Module1'});
   expect(result.ok).toBe(false);
  });
 });
}
