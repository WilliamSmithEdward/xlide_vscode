// Diagnostics tests: a For header continued over two lines keeps its range
// (issue #289). Measured in Excel 16.0 64-bit (2026-10-02) through
// pyVBAharness: each raises as the one-line header does.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function found(declaration: string, header: string, body: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    ${declaration}\n    ${header}\n        ${body}\n    Next\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a For header with a line continuation', () => {
	it('overflows an Integer counter as the one-line header does', () => {
		expect(found('Dim i As Integer', 'For i = _\n        1 To 40000', 'Main = i')).toEqual(['for-counter-overflow']);
	});

	it('keeps the counter range for the rules that read it', () => {
		expect(found('Dim a(1) As Long, i As Long', 'For i = _\n        1 To 3', 'a(i) = 1')).toEqual(['array-subscript-out-of-bounds']);
		expect(found('Dim a(1) As Long, i As Long', 'For i = 1 _\n        To 3', 'a(i) = 1')).toEqual(['array-subscript-out-of-bounds']);
	});

	it('stays quiet within range', () => {
		expect(found('Dim a(3) As Long, i As Long', 'For i = _\n        1 To 3', 'a(i) = 1')).toEqual([]);
		expect(found('Dim i As Integer', 'For i = _\n        1 To 30000', 'Main = i')).toEqual([]);
	});
});
