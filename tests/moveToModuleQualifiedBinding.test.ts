import { describe, expect, it } from 'vitest';
import { moveToModule } from '../src/analyzer/refactor/moveToModule';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';
for (const eol of ['\n', '\r\n', '\r']) describe('qualified module bindings ' + JSON.stringify(eol), () => {
 const cases = [
  {name: 'escaped receiver', body: '[Reports].Build', expected: '[Helpers].Build'},
  {name: 'escaped member', body: 'Reports.[Build]', expected: 'Helpers.[Build]'},
  {name: 'both escaped', body: '[Reports].[Build]', expected: '[Helpers].[Build]'},
  {name: 'continued receiver', body: 'Reports _' + eol + '.Build', expected: 'Helpers _' + eol + '.Build'},
  {name: 'continued member', body: 'Reports. _' + eol + 'Build', expected: 'Helpers. _' + eol + 'Build'},
  {name: 'parameter shadow', parameter: 'ByVal Reports As Object', body: 'Reports.Build', expected: 'Reports.Build'},
  {name: 'local shadow', declarations: 'Dim Reports As Object' + eol, body: 'Reports.Build', expected: 'Reports.Build'},
  {name: 'module field shadow', prefix: 'Private Reports As Object' + eol, body: 'Reports.Build', expected: 'Reports.Build'},
  {name: 'long receiver chain', parameter: 'ByVal item As Object', body: 'item.Reports.Build', expected: 'item.Reports.Build'},
  {name: 'bang receiver chain', parameter: 'ByVal item As Object', body: 'item!Reports.Build', expected: 'item!Reports.Build'},
  {name: 'module procedure pointer', body: 'Install AddressOf Reports.Build', expected: 'Install AddressOf Helpers.Build'},
  {name: 'ordinary qualified reference', body: 'Reports.Build', expected: 'Helpers.Build'},
 ];
 for (const test of cases) it(test.name, () => {
  const source = ['Public Sub Build()', 'End Sub', ''].join(eol);
  const caller = (test.prefix ?? '') + 'Sub Caller(' + (test.parameter ?? '') + ')' + eol + (test.declarations ?? '') + test.body + eol + 'Debug.Print "Reports.Build"' + eol + "' Reports.Build" + eol + 'End Sub' + eol;
  const expected = (test.prefix ?? '') + 'Sub Caller(' + (test.parameter ?? '') + ')' + eol + (test.declarations ?? '') + test.expected + eol + 'Debug.Print "Reports.Build"' + eol + "' Reports.Build" + eol + 'End Sub' + eol;
  const result = moveToModule({source, offset: source.indexOf('Build'), moduleName: 'Reports', targetModuleName: 'Helpers', otherModuleSources: {Helpers: '', Caller: caller}});
  if (!result.ok) throw new Error(result.reason);
  expect(applyVbaTextEdits(source, result.edits)).toBe('');
  const target = result.otherModules!.find(module => module.moduleName === 'Helpers')!;
  expect(applyVbaTextEdits('', target.edits)).toBe(eol + source);
  const edits = result.otherModules!.find(module => module.moduleName === 'Caller')?.edits ?? [];
  expect(applyVbaTextEdits(caller, edits)).toBe(expected);
  for (const edit of edits) { expect(edit.span.start).toBeGreaterThanOrEqual(0); expect(edit.span.end).toBeLessThanOrEqual(caller.length); }
  for (let i=1;i<edits.length;i++) expect(edits[i-1].span.end).toBeLessThanOrEqual(edits[i].span.start);
 });
});
