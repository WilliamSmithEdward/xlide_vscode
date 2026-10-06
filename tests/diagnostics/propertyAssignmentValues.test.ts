import { describe, expect, it } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';
import type { HostObjectModel } from '../../src/analyzer';

function excel(statement: string) {
 const source = `Option Explicit\nSub T(ByVal ws As Worksheet)\n${statement}\nEnd Sub`;
 return {source, diagnostics: analyzeVbaModuleSource({source,moduleName:'Module1',host:'excel',referencedHosts:[]}).diagnostics};
}
describe('property assignment diagnostics', () => {
 it.each(['ws.Range("A1").EntireRow.Height = 20', 'With ws.Range("A1").EntireRow\n.Height = 20\nEnd With', 'If True Then ws.Range("A1").Height = 20'])('flags getter-only height: %s', statement => {
  const {source, diagnostics} = excel(statement);
  const hits = diagnostics.filter(d=>d.code==='host-readonly-value-assignment');
  expect(hits).toHaveLength(1);
  expect(source.slice(hits[0].span.start,hits[0].span.end)).toBe('Height');
  expect(hits[0]).toMatchObject({severity:'error',origin:'run'});
  expect(hits[0].message).toContain('RowHeight');
  expect(hits[0].message).not.toContain('compile error');
 });
 it.each(['ws.Range("A1").EntireRow.RowHeight = 20', 'ws.Range("A1").EntireColumn.ColumnWidth = 20', 'Debug.Print ws.Range("A1").Height', 'If ws.Range("A1").Height = 20 Then Exit Sub', 'ws.UsedRange = 1', 'ws.Range("A1").Cells(1, 1) = 1'])('preserves valid reads, setters and default writes: %s', statement => {
  expect(excel(statement).diagnostics).toEqual([]);
 });
 it.each(['Height','Width','Left','Top','Text','CountLarge','HasArray','HasFormula'])('flags the read-only scalar Range.%s even though its typelib type is Variant', name => {
  expect(excel(`ws.Range("A1").${name} = 20`).diagnostics.map(d=>d.code)).toContain('host-readonly-value-assignment');
 });
 it.each(['ws.Range("A1").RowHeight = "abc"', 'ws.Range("A1").HorizontalAlignment = "abc"', 'ws.Range("A1").HorizontalAlignment = 123'])('flags a known invalid host property value: %s', statement => {
  expect(excel(statement).diagnostics).toEqual([expect.objectContaining({code:'host-property-value-out-of-range',severity:'error'})]);
 });
 it('flags the actual Boolean host setter without rejecting valid conversions', () => {
  expect(excel('ws.EnableCalculation = "nonsense"').diagnostics).toContainEqual(expect.objectContaining({code:'assignment-type-mismatch',severity:'error',origin:'run'}));
  expect(excel('ws.EnableCalculation = True').diagnostics).toEqual([]);
  expect(excel('ws.EnableCalculation = "True"').diagnostics).toEqual([]);
 });
 it('respects runtime error handling', () => {
  const {diagnostics}=excel('On Error Resume Next\nws.Range("A1").Height = 20');
  expect(diagnostics.filter(d=>d.code==='host-readonly-value-assignment')).toEqual([]);
 });
 const model:HostObjectModel={source:'setter coercion fixture',aliases:{widget:'Fixture.Widget'},globals:{},types:{'Fixture.Widget':{displayName:'Widget',members:[{name:'Enabled',kind:'property',declaredType:'Boolean',access:'read/write'},{name:'Size',kind:'property',declaredType:'Double',access:'read/write'}]}}};
 it.each(['Size = "nonsense"', 'Enabled = "nonsense"'])('flags a provably incompatible scalar host setter: %s', statement => {
  expect(analyzeModule(`Option Explicit\nSub T(ByVal item As Widget)\nitem.${statement}\nEnd Sub`,{hostModel:model})).toContainEqual(expect.objectContaining({code:'assignment-type-mismatch',severity:'error'}));
 });
 it.each(['Size = "20"', 'Size = 20', 'Enabled = "True"', 'Enabled = True'])('allows VBA scalar coercion: %s', statement => {
  expect(analyzeModule(`Option Explicit\nSub T(ByVal item As Widget)\nitem.${statement}\nEnd Sub`,{hostModel:model})).toEqual([]);
 });
});
