import { describe, expect, it } from 'vitest';
import { callSitesOf } from '../src/analyzer/refactor/callSites';
import { introduceParameter } from '../src/analyzer/refactor/introduceParameter';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';
const examples = [
 ['Go "hello"', 'Go "hello", 3'],
 ['Go "a""b:can\'t"', 'Go "a""b:can\'t", 3'],
 ['Go #12:30#', 'Go #12:30#, 3'],
 ['Go first:=1', 'Go first:=1, added:=3'],
 ['Go 1: Debug.Print "live"', 'Go 1, 3: Debug.Print "live"'],
 ['Go "hello" \'note', 'Go "hello", 3 \'note'],
 ['Go 1 Rem note', 'Go 1, 3 Rem note'],
 ['Go (1)', 'Go (1), 3'],
 ['Go ((1 + 2))', 'Go ((1 + 2)), 3'],
 ['Module1.Go (1): Debug.Print 2', 'Module1.Go (1), 3: Debug.Print 2'],
 ['Call Go(1)', 'Call Go(1, 3)'],
 ['x = Go(1)', 'x = Go(1, 3)'],
 ['Debug.Print Go(1)', 'Debug.Print Go(1, 3)'],
 ['If True Then Go (1) Else Debug.Print 2', 'If True Then Go (1), 3 Else Debug.Print 2'],
 ['Go \'note', 'Go 3 \'note'],
 ['Go: Go 1', 'Go: Go 1, 3'],
 ['Label: Go "hello"', 'Label: Go "hello", 3'],
];
for (const eol of ['\n','\r\n','\r']) {
 describe('call argument fidelity '+JSON.stringify(eol), () => {
  for (const [before,after] of examples) {
   it(before, () => {
    const source = ['Option Explicit','Sub Caller()',before,'Debug.Print "neighbor"','End Sub',''].join(eol);
    const sites = callSitesOf(source,'Go');
    expect(sites).toHaveLength(1);
    const edits = sites.map(site => ({span:site.argumentInsert,newText:site.argumentText('3', 'added')}));
    expect(applyVbaTextEdits(source,edits)).toBe(source.replace(before,after));
   });
  }
  it('does not close a call using parentheses in its comment', () => {
   const source = ['Sub Caller()',"Call Go(1 \'note)",'End Sub'].join(eol);
   expect(callSitesOf(source,'Go')).toEqual([]);
  });
  it('edits calls on a physical line shared with a procedure header', () => {
   const source = 'Sub Caller(): Go "hello": End Sub'+eol;
   const sites = callSitesOf(source,'Go');
   expect(sites).toHaveLength(1);
   expect(applyVbaTextEdits(source,sites.map(site=>({span:site.argumentInsert,newText:site.argumentText('3')}))))
    .toBe(source.replace('Go "hello"','Go "hello", 3'));
  });
  it('preserves string arguments in actual same-module refactoring', () => {
   const source = ['Option Explicit','Public Sub Report(ByVal title As String)','Dim limit As Long','limit = 3','Debug.Print title, limit','End Sub','Sub Caller()', 'Report "hello": Debug.Print "live"', 'Call Report("other")', 'End Sub',''].join(eol);
   const result = introduceParameter({source,offset:source.indexOf('limit'),moduleName:'Module1'});
   if(!result.ok)throw new Error(result.reason);
   expect(applyVbaTextEdits(source,result.edits)).toBe(['Option Explicit','Public Sub Report(ByVal title As String, ByVal limit As Long)','Debug.Print title, limit','End Sub','Sub Caller()', 'Report "hello", 3: Debug.Print "live"', 'Call Report("other", 3)', 'End Sub',''].join(eol));
  });
  it('preserves qualified calls in actual external-module refactoring', () => {
   const source = ['Public Sub Report(ByVal title As String)','Dim limit As Long','limit = 3','Debug.Print title, limit','End Sub',''].join(eol);
   const other = ['Sub Caller()', 'Module1.Report "hello" \'note','Debug.Print "live"','End Sub',''].join(eol);
   const result = introduceParameter({source,offset:source.indexOf('limit'),moduleName:'Module1',otherModuleSources:{Other:other}});
   if(!result.ok)throw new Error(result.reason);
   expect(applyVbaTextEdits(other,result.otherModules![0].edits)).toBe(other.replace('Report "hello"','Report "hello", 3'));
  });
 });
}
