// Diagnostics tests: a For Each control variable whose type does not fit
// what the collection hands back (issue #447). Measured in Excel 16.0 (build
// 20326, 2026-10-02).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

const CODE = 'assignment-object-type-mismatch';

function source(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\n    Main = 1\nEnd Function\n`;
}

describe('For Each into a control variable of the wrong type', () => {
	it('raises 13 for a host item of another class', () => {
		for (const [lines, span] of [
			[['Dim c As Range', 'For Each c In Worksheets: Next'], 'Worksheets'],
			[['Dim w As Worksheet', 'For Each w In Range("A1:A2"): Next'], 'Range("A1:A2")'],
			[['Dim w As Worksheet', 'For Each w In Workbooks: Next'], 'Workbooks'],
		] as const) {
			const src = source(...lines);
			expectDiagnostic(src, byCode(analyzeModule(src), CODE), CODE, { span, message: "'13'" });
		}
	});

	it('raises 424 for a number or string into an object variable', () => {
		for (const [decl, add] of [['Collection', 'col.Add 1'], ['Object', 'col.Add 1'], ['Collection', 'col.Add "a"']]) {
			const src = source(`Dim o As ${decl}, col As New Collection`, add, 'For Each o In col: Next');
			expectDiagnostic(src, byCode(analyzeModule(src), CODE), CODE, { span: 'col', message: "'424'" });
		}
	});

	it('runs for the items each collection hands back', () => {
		for (const lines of [
			['Dim w As Worksheet', 'For Each w In Worksheets: Next'],
			['Dim w As Worksheet', 'For Each w In ThisWorkbook.Worksheets: Next'],
			['Dim r As Range', 'For Each r In Range("A1:B2"): Next'],
			['Dim r As Range', 'For Each r In Range("A1:B2").Rows: Next'],
			['Dim r As Range', 'For Each r In Range("A1:B2").Areas: Next'],
			['Dim b As Workbook', 'For Each b In Workbooks: Next'],
			['Dim n As Name', 'For Each n In ThisWorkbook.Names: Next'],
			['Dim s As Shape', 'For Each s In ActiveSheet.Shapes: Next'],
			['Dim v As Variant, col As New Collection', 'col.Add 1', 'For Each v In col: Next'],
			['Dim w As Worksheet', 'For Each w In Sheets: Next'],
			['Dim o As Object', 'For Each o In Worksheets: Next'],
		]) {
			expect(byCode(analyzeModule(source(...lines)), CODE), lines.join(' : ')).toEqual([]);
		}
	});
});
