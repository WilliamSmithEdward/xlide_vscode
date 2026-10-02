// Diagnostics tests: Excel's Worksheets and Charts properties return a Sheets
// object, so Set into a variable As Worksheets or As Charts raises 13 (issue
// #404). Measured in Excel 16.0 (build 20326, 2026-10-01).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

const CODE = 'assignment-object-type-mismatch';

function source(lines: string[], after = ''): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\n    Main = 1\nEnd Function\n${after}`;
}

describe('Set from Worksheets or Charts', () => {
	it('raises 13 into a variable As Worksheets or As Charts', () => {
		for (const [lines, span] of [
			[['Dim w As Worksheets', 'Set w = Worksheets'], 'Worksheets'],
			[['Dim w As Worksheets', 'Set w = ThisWorkbook.Worksheets'], 'ThisWorkbook.Worksheets'],
			[['Dim w As Worksheets', 'Set w = ActiveWorkbook.Worksheets'], 'ActiveWorkbook.Worksheets'],
			[['Dim w As Worksheets', 'Set w = Worksheets(Array("Sheet1"))'], 'Worksheets(Array("Sheet1"))'],
			[['Dim w As Excel.Worksheets', 'Set w = Worksheets'], 'Worksheets'],
			[['Dim c As Charts', 'Set c = Charts'], 'Charts'],
			[['Dim c As Charts', 'Set c = ThisWorkbook.Charts'], 'ThisWorkbook.Charts'],
		] as const) {
			const src = source([...lines]);
			expectDiagnostic(src, byCode(analyzeModule(src), CODE), CODE, { span, message: ['Sheets object', "'13'"] });
		}
	});

	it('runs into a variable As Sheets, and for the other collections', () => {
		for (const lines of [
			['Dim s As Sheets', 'Set s = Worksheets'],
			['Dim s As Sheets', 'Set s = Charts'],
			['Dim n As Names', 'Set n = Names'],
			['Dim b As Workbooks', 'Set b = Workbooks'],
			['Dim w As Worksheet', 'Set w = Worksheets(1)'],
		]) {
			expect(byCode(analyzeModule(source(lines)), CODE), lines.join(' : ')).toEqual([]);
		}
	});

	it('leaves a variable named Worksheets alone', () => {
		const src = source(['Dim w As Worksheets, Worksheets As Worksheets', 'Set w = Worksheets']);
		expect(byCode(analyzeModule(src), CODE)).toEqual([]);
	});

	it('raises 13 passed to a parameter As Worksheets', () => {
		const src = source(['Main = TakeW(Worksheets)'], 'Private Function TakeW(w As Worksheets) As Long\nEnd Function\n');
		expectDiagnostic(src, byCode(analyzeModule(src), 'argument-type-mismatch'), 'argument-type-mismatch', { span: 'Worksheets', message: ['Sheets object', "'13'"] });
		const sheets = source(['Main = TakeS(Worksheets)'], 'Private Function TakeS(s As Sheets) As Long\nEnd Function\n');
		expect(byCode(analyzeModule(sheets), 'argument-type-mismatch')).toEqual([]);
	});

	it('raises 13 as a Function As Worksheets returns it', () => {
		const src = source(['Main = 1'], 'Private Function GetW() As Worksheets\n    Set GetW = Worksheets\nEnd Function\n');
		expectDiagnostic(src, byCode(analyzeModule(src), CODE), CODE, { span: 'Worksheets', message: 'Sheets object' });
	});
});
