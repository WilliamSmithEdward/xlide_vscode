import { describe, expect, it } from 'vitest';
import { analyzeModule } from '../src/analyzer';
import { analyzeProjectModule } from './diagnostics/helpers';
const fn = 'Function MakeFlags() As Boolean()\nDim flags(1) As Boolean\nMakeFlags = flags\nEnd Function\n';
const sourceOf = (line: string) => 'Option Explicit\n' + fn + 'Function Main() As Variant\nDim values(1) As Boolean\n' + line + '\nMain = 1\nEnd Function';
describe('reported runtime assignment gaps', () => {
 it.each(['Err.LastDllError', 'VBA.Err.LastDllError'])('rejects read-only %s', target => {
  const source = sourceOf(target + ' = 5');
  for (const pass of [analyzeModule(source), analyzeProjectModule(source, [], 'Caller')]) {
   const found = pass.filter(d => d.code === 'readonly-member-assignment');
   expect(found).toHaveLength(1);
   expect(found[0].message).toContain("Can't assign to read-only property");
   expect(source.slice(found[0].span.start, found[0].span.end)).toBe('LastDllError');
  }
 });
 it.each(['values', '[values]', '(values)', 'MakeFlags()', '(MakeFlags())'])('rejects whole-array %s at the workbook Worksheet setter', value => {
  const source = sourceOf('ThisWorkbook.Worksheets(1).EnableCalculation = ' + value);
  for (const found of [analyzeModule(source), analyzeProjectModule(source, [], 'Caller')]) {
   const errors = found.filter(d => d.code === 'assignment-type-mismatch');
   expect(errors).toHaveLength(1);
   expect(errors[0].message).toContain("Run-time error '13'");
   expect(errors[0].message).not.toContain('compile error');
  }
 });
 it.each(['Err.Number = 5', 'VBA.Err.Number = 5', 'Main = Err.LastDllError'])('preserves valid runtime member use: %s', line => {
  expect(analyzeModule(sourceOf(line)).filter(d => d.severity === 'error')).toEqual([]);
 });
 it.each(['True', 'values(0)', '"True"'])('preserves valid scalar host values: %s', value => {
  expect(analyzeModule(sourceOf('ThisWorkbook.Worksheets(1).EnableCalculation = ' + value)).filter(d => d.severity === 'error')).toEqual([]);
 });
 it.each(['ThisWorkbook.Sheets(1)', 'unknown'])('does not assume a Boolean setter for ambiguous %s', receiver => {
  const source = sourceOf(receiver + '.EnableCalculation = values').replace('Dim values(1) As Boolean', 'Dim values(1) As Boolean\nDim unknown As Object');
  expect(analyzeModule(source).filter(d => d.code === 'assignment-type-mismatch')).toEqual([]);
 });
 it('does not confuse a shadowed Err with the runtime object', () => {
  const source = 'Sub T(ByVal Err As Object)\nErr.LastDllError = 5\nEnd Sub';
  expect(analyzeModule(source).filter(d => d.code === 'readonly-member-assignment')).toEqual([]);
 });
 it('retains compile-time readonly checking under runtime error handling', () => {
  expect(analyzeModule(sourceOf('On Error Resume Next\nErr.LastDllError = 5')).filter(d => d.code === 'readonly-member-assignment')).toHaveLength(1);
 });
 it('resolves the host contract inside With and a single-line If', () => {
  for (const line of ['With ThisWorkbook.Worksheets(1)\n.EnableCalculation = values\nEnd With', 'If True Then ThisWorkbook.Worksheets(1).EnableCalculation = MakeFlags()']) {
   expect(analyzeModule(sourceOf(line)).filter(d => d.code === 'assignment-type-mismatch')).toHaveLength(1);
  }
 });
});

