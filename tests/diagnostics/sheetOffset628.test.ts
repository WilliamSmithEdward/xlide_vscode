// Diagnostics tests: Offset past the sheet's edge on a whole row, a whole
// column or the whole sheet, named by Rows, Columns, Range("5:5") or Cells
// (issue #628). Measured on 2026-10-02 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(body: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    ${body}\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('Offset past the edge of a whole row, column or sheet (issue #628)', () => {
	it('reports a step that leaves the sheet', () => {
		for (const body of [
			'Main = ActiveSheet.Rows(5).Offset(0, 1).Address',
			'Main = ActiveSheet.Columns(3).Offset(1, 0).Address',
			'Main = ActiveSheet.Range("5:5").Offset(0, 1).Address',
			'Main = ActiveSheet.Range("C:C").Offset(1, 0).Address',
			'Main = ActiveSheet.Cells.Offset(1, 0).Address',
			'Main = ActiveSheet.Rows.Offset(1, 0).Address',
			'Main = Rows.Offset(0, 1).Address',
			'Main = Cells.Offset(0, 1).Address',
			'Main = Columns.Offset(1, 0).Address',
			'Main = Worksheets(1).Rows(5).Offset(0, 1).Address',
			'Main = Sheets(1).Columns(3).Offset(1, 0).Address',
			'Main = Application.ActiveSheet.Rows(5).Offset(0, 1).Address',
			'Main = ActiveSheet.Range("5:5").Offset(-5, 0).Address',
			'Main = ActiveSheet.Rows(1).Offset(-1, 0).Address',
			'Main = ActiveSheet.Range("5:6").Offset(0, 1).Address',
			'Main = ActiveSheet.Range("C:D").Offset(1, 0).Address',
			'Main = ActiveSheet.Range("$5:$5").Offset(0, 1).Address',
		]) {
			expect(errors(body), body).toEqual(['host-argument-out-of-range']);
		}
	});

	it('stays quiet where the block stays on the sheet', () => {
		for (const body of [
			'Main = ActiveSheet.Rows(5).Offset(1, 0).Address',
			'Main = ActiveSheet.Columns(3).Offset(0, 1).Address',
			'Main = ActiveSheet.Range("5:5").Offset(1, 0).Address',
			'Main = ActiveSheet.Range("C:C").Offset(0, 1).Address',
			'Main = ActiveSheet.Cells.Offset(0, 0).Address',
			'Main = ActiveSheet.Range("5:5").Offset(-4, 0).Address',
			'Main = ActiveSheet.Range("A1").Rows(5).Offset(0, 1).Address',
			'Main = ActiveSheet.Rows(5).Item(2).Address',
			'Main = ActiveSheet.Range("5:5").Item(2).Address',
			'Main = ActiveSheet.Range("5:5").Resize(1, 2).Address',
			'Main = ActiveSheet.Rows(5).Resize(2).Address',
			// A local of that name is the code's own.
			'Dim Rows As Range\n    Set Rows = Range("A1")\n    Main = Rows.Offset(1, 0).Address',
		]) {
			expect(errors(body), body).toEqual([]);
		}
	});
});
