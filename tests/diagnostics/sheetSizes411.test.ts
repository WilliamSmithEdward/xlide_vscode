// Diagnostics tests: sheet sizes Excel fixes, read through a qualified
// receiver or a literal address, into an Integer or Byte (issue #411).
// Measured in Excel 16.0 64-bit (build 20430, 2026-10-02) through
// pyVBAharness.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(...lines: string[]): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    Dim i As Integer, b As Byte, l As Long, ws As Worksheet\n    Set ws = ActiveSheet\n${lines.map((line) => `    ${line}`).join('\n')}\n    Main = 1\nEnd Function\n`;
	return analyzeModule(src, { host: 'excel' }).filter((diag) => diag.severity === 'error').map((diag) => `${diag.code}: ${diag.message}`);
}

describe('a sheet size into an Integer', () => {
	it('overflows through a qualified receiver or a literal address', () => {
		for (const lines of [
			['i = ActiveSheet.Rows.Count'], ['i = Worksheets(1).Rows.Count'], ['i = ws.Rows.Count'], ['i = Application.Rows.Count'],
			['i = ActiveSheet.Rows.Count \\ 2'], ['For i = 1 To ws.Rows.Count', 'Exit For', 'Next'],
			['For i = 1 To ThisWorkbook.Worksheets(1).Rows.Count', 'Exit For', 'Next'], ['b = ws.Columns.Count'],
			['i = Cells(Rows.Count, 1).Row'], ['i = ws.Cells(ws.Rows.Count, 1).Row'], ['i = Range("A40000").Row'],
			['i = Range("A1:A40000").Rows.Count'], ['i = Range("A1:B20000").Count'], ['i = Range("A:A").Count'],
		]) {
			expect(errors(...lines), lines.join(' / ')).toEqual([expect.stringContaining("'6'")]);
		}
	});

	it('is quiet where the size fits', () => {
		for (const line of [
			'i = ws.Columns.Count \\ 2', 'l = ws.Columns.Count', 'i = Range("A1:A30000").Rows.Count', 'i = Range("A1:B2").Count',
			'i = Cells(5, 1).Row', 'i = Range("C5").Column', 'i = Range("1:1").Count', 'i = Range("B2:D9").Columns.Count',
			'l = ActiveSheet.Columns.Count',
		]) {
			expect(errors(line), line).toEqual([]);
		}
	});
});
