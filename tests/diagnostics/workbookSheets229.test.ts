// Saved sheet metadata must never be treated as runtime collection contents.
import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { mergeSheetChanges, sheetChangesIn, type WorkbookSheetInfo } from '../../src/analyzer/symbols/sheetChanges';
import { byCode } from '../helpers/diagnostics';

const SHEETS: WorkbookSheetInfo[] = [
 { name: 'Budget', kind: 'worksheet' },
 { name: 'Trend', kind: 'chartsheet' },
];
function wrap(...lines: string[]): string {
 return 'Option Explicit\nFunction Main() As Variant\n' + lines.join('\n') + '\nEnd Function\n';
}
function hits(source: string, sheets: WorkbookSheetInfo[] = SHEETS, others: string[] = []) {
 return byCode(analyzeModule(source, {
  workbookSheets: sheets,
  projectSheetChanges: mergeSheetChanges([source, ...others].map(sheetChangesIn)),
 }), 'sheet-not-in-workbook');
}

describe('saved sheet contents are not runtime facts', () => {
 it.each([
  'Main = ThisWorkbook.Sheets("Report").Name',
  'Main = ThisWorkbook.Sheets(10).Name',
  'Main = ThisWorkbook.Worksheets("Trend").Name',
  'Main = ThisWorkbook.Charts("Budget").Name',
  'Main = ThisWorkbook.Sheets.Item("Nope").Name',
  'Main = Application.ThisWorkbook.Worksheets("Nope").Name',
  'Main = ThisWorkbook.Sheets("1").Name',
  'Dim i As Long\nFor i = 1 To 10\nMain = ThisWorkbook.Sheets(i).Name\nNext',
 ])('leaves snapshot mismatches alone: %s', line => {
  expect(hits(wrap(line))).toEqual([]);
 });
 it('leaves Charts(1) alone even when the saved workbook has no charts', () => {
  const sheets: WorkbookSheetInfo[] = [{ name: 'Budget', kind: 'worksheet' }];
  expect(hits(wrap('ThisWorkbook.Charts.Add2', 'ThisWorkbook.Charts(1).Activate'), sheets)).toEqual([]);
  // No local mutation need be visible: another macro or the user may add it.
  expect(hits(wrap('ThisWorkbook.Charts(1).Activate'), sheets)).toEqual([]);
 });
 it.each([
  'Dim wb As Workbook\nSet wb = ThisWorkbook\nwb.Worksheets.Add.Name = "Report"',
  'ThisWorkbook.Sheets("Budget").Name = "Report"',
  'ThisWorkbook.Sheets("Budget").Delete',
  'Application.Run "BuildReport"',
  'DoEvents',
 ])('stays quiet with runtime changes: %s', change => {
  expect(hits(wrap(change, 'Main = ThisWorkbook.Sheets("Report").Name', 'Main = ThisWorkbook.Sheets(10).Name'))).toEqual([]);
 });
 it('keeps snapshot-independent invalid argument and name checks', () => {
  const source = wrap('Main = ThisWorkbook.Sheets(0).Name', 'ThisWorkbook.Worksheets(1).Name = "a:b"');
  const diagnostics = analyzeModule(source, { workbookSheets: SHEETS, projectSheetChanges: sheetChangesIn(source) });
  expect(byCode(diagnostics, 'host-argument-out-of-range')).toHaveLength(1);
  expect(byCode(diagnostics, 'sheet-name-invalid')).toHaveLength(1);
 });
});

describe('sheetChangesIn (issue #229)', () => {
	it('names what a module does to sheets', () => {
		const changes = sheetChangesIn('Sub A(s)\n    Worksheets.Add.Name = "Out"\n    If s = 1 Then Sheet2.Name = "Two" Else Sheet3.Name = "Three"\nEnd Sub\n');
		expect(changes.addsSheets).toBe(true);
		expect([...changes.namesAssigned].sort()).toEqual(['out', 'three', 'two']);
		expect(changes.assignsComputedName).toBe(false);
	});
});
