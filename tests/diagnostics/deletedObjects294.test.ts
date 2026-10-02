// Diagnostics tests: an object used after Delete, Close or Unlist, and what
// Erase leaves (issue #294). Measured in Excel 16.0 64-bit (2026-10-02)
// through pyVBAharness.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function found(body: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    ${body}\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

const ADD = 'Dim ws As Worksheet\n    Set ws = Worksheets.Add\n    ';

describe('an object after the statement that ends it', () => {
	it('raises on a member', () => {
		expect(found(`${ADD}ws.Delete\n    Main = ws.Name`)).toEqual(['object-used-after-delete']);
		expect(found(`${ADD}Dim r As Range\n    Set r = ws.Range("A1")\n    ws.Delete\n    Main = r.Address`)).toEqual(['object-used-after-delete']);
		expect(found('Dim wb As Workbook\n    Set wb = Workbooks.Add\n    wb.Close False\n    Main = wb.Name')).toEqual(['object-used-after-delete']);
		expect(found('Dim n As Name\n    Set n = ThisWorkbook.Names.Add("xlideNameOne", "=1")\n    n.Delete\n    Main = n.Name')).toEqual(['object-used-after-delete']);
	});

	it('stays quiet on Is Nothing, a new Set, and a use before the end', () => {
		expect(found(`${ADD}ws.Delete\n    Main = ws Is Nothing`)).toEqual([]);
		expect(found(`${ADD}ws.Delete\n    Set ws = Worksheets(1)\n    Main = ws.Name`)).toEqual([]);
		expect(found('Dim n As Name\n    Set n = ThisWorkbook.Names.Add("xlideNameTwo", "=1")\n    Main = n.Name\n    n.Delete')).toEqual([]);
	});
});

describe('Erase', () => {
	it('leaves a fixed array of objects Nothing and a Variant array empty', () => {
		expect(found('Dim o(1) As Collection\n    Set o(0) = New Collection\n    Erase o\n    Main = o(0).Count')).toEqual(['object-variable-not-set']);
		expect(found('Dim v As Variant\n    v = Array(1, 2)\n    Erase v\n    Main = v(0)')).toEqual(['unallocated-dynamic-array-access']);
	});
});
