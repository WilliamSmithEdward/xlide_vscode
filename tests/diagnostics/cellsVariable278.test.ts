// Diagnostics tests: a Range local that one Set gives a sheet's Cells counts
// past the Long range (issue #278). Each case was run through pyVBAharness on
// 2026-10-02 in Excel 16.0 64-bit (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(body: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('Set r = Cells (issue #278)', () => {
	it('counts the whole sheet through the variable', () => {
		for (const source of ['Cells', 'ActiveSheet.Cells', 'Worksheets(1).Cells']) {
			expect(errors(`Dim r As Range\n    Set r = ${source}\n    Main = r.Count`), source).toEqual(['arithmetic-overflow']);
		}
		expect(errors('Dim r As Object\n    Set r = Cells\n    Main = r.Count')).toEqual(['arithmetic-overflow']);
		expect(errors('Dim r As Range, n As Long\n    Set r = Cells\n    n = r.Cells.Count')).toEqual(['arithmetic-overflow']);
	});

	it('stays quiet for CountLarge, a row count, a block, and a variable set again', () => {
		expect(errors('Dim r As Range\n    Set r = Cells\n    Main = r.CountLarge')).toEqual([]);
		expect(errors('Dim r As Range\n    Set r = Cells\n    Main = r.Rows.Count')).toEqual([]);
		expect(errors('Dim r As Range\n    Set r = Range("A1:B2")\n    Main = r.Count')).toEqual([]);
		expect(errors('Dim r As Range\n    Set r = Cells\n    Set r = Range("A1")\n    Main = r.Count')).toEqual([]);
	});
});
