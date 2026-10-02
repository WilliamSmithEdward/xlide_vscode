// Diagnostics tests: Cells.Count past the Long range, and the shape of a
// range's value (issue #278). Measured in Excel 16.0 64-bit (2026-10-02)
// through pyVBAharness.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function found(body: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    ${body}\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a whole sheet counted cell by cell', () => {
	it('overflows Range.Count', () => {
		for (const body of ['Main = Cells.Count', 'Main = ActiveSheet.Cells.Count', 'Main = Worksheets(1).Cells.Count', 'Main = Range("A:XFD").Count', 'If Cells.Count > 1 Then Main = 1']) {
			expect(found(body), body).toEqual(['arithmetic-overflow']);
		}
	});

	it('stays quiet on CountLarge, rows, columns, smaller blocks and a variable named cells', () => {
		for (const body of ['Main = Cells.CountLarge', 'Main = Rows.Count', 'Main = Cells.Rows.Count', 'Main = Range("A:C").Count', 'Main = Range("1:3000").Count', 'Dim cells As New Collection\n    Main = cells.Count']) {
			expect(found(body), body).toEqual([]);
		}
	});
});

describe('the value of a range', () => {
	it('is no array for one cell', () => {
		expect(found('Dim v As Variant\n    v = Range("A1").Value\n    Main = v(1, 1)')).toEqual(['variant-value-misuse']);
		expect(found('Dim v As Variant\n    v = Range("A1").Value\n    Main = UBound(v)')).toEqual(['variant-value-misuse']);
		expect(found('Dim v As Variant\n    v = Cells(1, 1).Value\n    Main = v(1, 1)')).toEqual(['variant-value-misuse']);
		expect(found('Dim v As Variant, x As Variant\n    v = Range("A1").Value\n    For Each x In v\n    Next')).toEqual(['variant-value-misuse']);
	});

	it('is a 2-D array from 1 for a block, through Value2 too, and Transpose turns it', () => {
		expect(found('Dim v As Variant\n    v = Range("A1:B2").Value2\n    Main = v(0, 0)')).toEqual(['array-subscript-out-of-bounds']);
		expect(found('Dim t As Variant\n    t = Application.Transpose(Range("A1:A3").Value)\n    Main = t(0)')).toEqual(['array-subscript-out-of-bounds']);
		expect(found('Dim t As Variant\n    t = Application.Transpose(Range("A1:A3").Value)\n    Main = t(1)')).toEqual([]);
	});

	it('has one area for an address with no comma', () => {
		expect(found('Main = Range("A1:B2").Areas(2).Address')).toEqual(['host-argument-out-of-range']);
		expect(found('Main = Range("A1,C3").Areas(2).Address')).toEqual([]);
	});
});
