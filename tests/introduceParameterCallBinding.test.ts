import { describe, expect, it } from 'vitest';
import { introduceParameter } from '../src/analyzer/refactor/introduceParameter';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';

for (const eol of ['\n', '\r\n', '\r']) {
 describe('Introduce Parameter call binding '+JSON.stringify(eol), () => {
  const target = ['Public Sub Report()', 'Dim limit As Long', 'limit = 3', 'Debug.Print limit', 'End Sub', ''].join(eol);
  const changed = ['Public Sub Report(ByVal limit As Long)', 'Debug.Print limit', 'End Sub', ''].join(eol);
  function outputs(source: string, others: Record<string,string>) {
   const result = introduceParameter({source, offset: source.indexOf('limit'), moduleName: 'Module1', otherModuleSources: others});
   if (!result.ok) { throw new Error(result.reason); }
   const external = {...others};
   for (const module of result.otherModules ?? []) { external[module.moduleName] = applyVbaTextEdits(others[module.moduleName], module.edits); }
   return {source: applyVbaTextEdits(source,result.edits), external};
  }
  it('updates only the owning module in the reported reproduction', () => {
   const caller = ['Sub Caller()', 'Module1.Report', 'Module2.Report', 'End Sub', ''].join(eol);
   const other = ['Public Sub Report()', 'Debug.Print "other"', 'End Sub', ''].join(eol);
   expect(outputs(target,{Caller:caller, Module2:other})).toEqual({source:changed,external:{Caller:caller.replace('Module1.Report','Module1.Report 3'),Module2:other}});
  });
  it('supports bracketed names and case-insensitive qualifiers', () => {
   const caller = ['Sub Caller()', 'Call [mOdUlE1].[Report]()', '[Module2].[Report]', 'End Sub', ''].join(eol);
   expect(outputs(target,{Caller:caller})).toEqual({source:changed,external:{Caller:caller.replace('[Report]()','[Report](3)')}});
  });
  it('binds unqualified calls to the same-module procedure before the project', () => {
   const other = ['Public Sub Report()', 'End Sub', 'Sub Caller()', 'Report', 'Module1.Report', 'End Sub', ''].join(eol);
   expect(outputs(target,{Module2:other})).toEqual({source:changed,external:{Module2:other.replace('Module1.Report','Module1.Report 3')}});
  });
  it('does not rewrite locals, parameters, return variables or unrelated members', () => {
   const caller = ['Sub LocalCaller()', 'Dim Report As Long', 'Report = 2', 'Debug.Print Report', 'End Sub', 'Sub ParamCaller(ByVal Report As Long)', 'Debug.Print Report', 'End Sub', 'Function Report() As Long', 'Report = 4', 'Debug.Print Report', 'End Function', 'Sub MemberCaller()', 'Dim item As Object', 'item.Report', 'With item', '.Report', 'End With', 'End Sub', ''].join(eol);
   expect(outputs(target,{Caller:caller})).toEqual({source:changed,external:{Caller:caller}});
  });
  it('does not treat a shadowed module name or a longer member chain as the module', () => {
   const caller = ['Sub Caller(ByVal Module1 As Object)', 'Module1.Report', 'End Sub', 'Sub Other()', 'Dim item As Object', 'item.Module1.Report', 'End Sub', ''].join(eol);
   expect(outputs(target,{Caller:caller})).toEqual({source:changed,external:{Caller:caller}});
  });
  it('filters qualified and shadowed calls in other procedures of the target module', () => {
   const local = ['Sub Caller()', 'Report', 'Module1.Report', 'Module2.Report', 'End Sub', 'Sub Shadowed(ByVal Report As Long)', 'Debug.Print Report', 'End Sub', ''].join(eol);
   const expected = ['Sub Caller()', 'Report 3', 'Module1.Report 3', 'Module2.Report', 'End Sub', 'Sub Shadowed(ByVal Report As Long)', 'Debug.Print Report', 'End Sub', ''].join(eol);
   expect(outputs(target+local,{})).toEqual({source:changed+expected,external:{}});
  });
  it('keeps ambiguous project calls unchanged but updates explicit owning-module calls', () => {
   const caller = ['Sub Caller()', 'Report', 'Module1.Report', 'End Sub', ''].join(eol);
   const other = ['Public Sub Report()', 'End Sub', ''].join(eol);
   expect(outputs(target,{Caller:caller,Module2:other})).toEqual({source:changed,external:{Caller:caller.replace('Module1.Report','Module1.Report 3'),Module2:other}});
  });
  it('updates unqualified calls when the target is the only visible declaration', () => {
   const caller = ['Sub Caller()', 'Report', 'Call Module1.Report()', 'End Sub', ''].join(eol);
   expect(outputs(target,{Caller:caller})).toEqual({source:changed,external:{Caller:['Sub Caller()', 'Report 3', 'Call Module1.Report(3)', 'End Sub', ''].join(eol)}});
  });
  it('preserves module variables and qualifiers shadowed at module scope', () => {
   const caller = ['Private Report As Long', 'Private Module1 As Object', 'Sub Caller()', 'Debug.Print Report', 'Module1.Report', 'End Sub', ''].join(eol);
   expect(outputs(target,{Caller:caller})).toEqual({source:changed,external:{Caller:caller}});
  });
  it('preserves labels, named-argument keys, comments and string literals', () => {
   const caller = ['Sub Caller()', 'Report:', 'GoTo Report', 'GoSub Report', 'Resume Report', 'Other Report:=2', 'Other AddressOf Module1.Report', 'Debug.Print "Report", "Module1.Report"', "' Module1.Report", 'Module1.Report', 'End Sub', ''].join(eol);
   expect(outputs(target,{Caller:caller})).toEqual({source:changed,external:{Caller:caller.replace(eol+'Module1.Report'+eol,eol+'Module1.Report 3'+eol)}});
  });
  it('binds continued qualified calls and preserves their argument text', () => {
   const caller = ['Sub Caller()', 'Call Module1. _', 'Report()', 'Module2. _', 'Report', 'End Sub', ''].join(eol);
   expect(outputs(target,{Caller:caller})).toEqual({source:changed,external:{Caller:caller.replace('Report()','Report(3)')}});
  });
  it('does not expose a private target to callers in another module', () => {
   const caller = ['Sub Caller()', 'Report', 'Module1.Report', 'End Sub', ''].join(eol);
   expect(outputs(target.replace('Public','Private'),{Caller:caller})).toEqual({source:changed.replace('Public','Private'),external:{Caller:caller}});
  });
  it('binds Function expressions while preserving another Function return variable', () => {
   const fn = ['Public Function Report() As Long', 'Dim limit As Long', 'limit = 3', 'Report = limit', 'End Function', ''].join(eol);
   const caller = ['Function Caller() As Long', 'Caller = Report + Module1.Report()', 'End Function', ''].join(eol);
   expect(outputs(fn,{Caller:caller})).toEqual({source:['Public Function Report(ByVal limit As Long) As Long', 'Report = limit', 'End Function', ''].join(eol),external:{Caller:['Function Caller() As Long', 'Caller = Report(3) + Module1.Report(3)', 'End Function', ''].join(eol)}});
  });

  it('distinguishes a conditional call before a colon from a statement label', () => {
   const caller = ['Sub Caller()', 'If True Then Report: Debug.Print 1', 'End Sub', ''].join(eol);
   expect(outputs(target,{Caller:caller})).toEqual({source:changed,external:{Caller:caller.replace('Then Report:', 'Then Report 3:')}});
  });
 });
}
