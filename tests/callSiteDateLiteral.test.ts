import { describe, expect, it } from 'vitest';
import { callSitesOf } from '../src/analyzer/refactor/callSites';
import { introduceParameter } from '../src/analyzer/refactor/introduceParameter';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';

for (const eol of ['\n', '\r\n', '\r']) {
 describe('call sites preserve date literals '+JSON.stringify(eol), () => {
  for (const month of ['January','February','March','April','May','June','July','August','September','October','November','December']) {
   it('preserves '+month+' inside a date and updates the real call', () => {
    const source=['Sub Caller()', 'Debug.Print #'+month+' 1, 2000#', month, 'End Sub', ''].join(eol);
    const sites=callSitesOf(source,month);
    expect(sites).toHaveLength(1);
    const output=applyVbaTextEdits(source,sites.map(site=>({span:site.argumentInsert,newText:site.argumentText('3')})));
    expect(output).toBe(['Sub Caller()', 'Debug.Print #'+month+' 1, 2000#', month+' 3', 'End Sub', ''].join(eol));
   });
  }
  it('preserves a date with a real call on the same physical line', () => {
   const source=['Sub Caller()', 'Debug.Print #May 1, 2000#: May', 'End Sub', ''].join(eol);
   const sites=callSitesOf(source,'May');
   expect(sites).toHaveLength(1);
   expect(applyVbaTextEdits(source,sites.map(site=>({span:site.argumentInsert,newText:site.argumentText('3')}))))
    .toBe(source.replace(': May', ': May 3'));
  });
  it('retains file-number markers and calls after them', () => {
   const source=['Sub Caller()', 'Print #channel, May(1)', 'Debug.Print #5/1/2000#', 'End Sub', ''].join(eol);
   const sites=callSitesOf(source,'May');
   expect(sites).toHaveLength(1);
   expect(applyVbaTextEdits(source,sites.map(site=>({span:site.argumentInsert,newText:site.argumentText('3')}))))
    .toBe(source.replace('May(1)', 'May(1, 3)'));
  });
  it('preserves dates in complete Introduce Parameter output', () => {
   const source=['Public Sub May()', 'Dim limit As Long', 'limit = 3', 'Debug.Print limit', 'End Sub', 'Sub Caller()', 'Debug.Print #May 1, 2000#', 'May', 'End Sub', ''].join(eol);
   const external=['Sub ExternalCaller()', 'Debug.Print #May 2, 2000#', 'Module1.May', 'End Sub', ''].join(eol);
   const result=introduceParameter({source,offset:source.indexOf('limit'),moduleName:'Module1',otherModuleSources:{Other:external}});
   if(!result.ok)throw new Error(result.reason);
   expect(applyVbaTextEdits(source,result.edits)).toBe(['Public Sub May(ByVal limit As Long)', 'Debug.Print limit', 'End Sub', 'Sub Caller()', 'Debug.Print #May 1, 2000#', 'May 3', 'End Sub', ''].join(eol));
   expect(result.otherModules).toHaveLength(1);
   expect(applyVbaTextEdits(external,result.otherModules![0].edits)).toBe(external.replace('Module1.May', 'Module1.May 3'));
  });
 });
}
