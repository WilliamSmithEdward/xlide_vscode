// Diagnostics tests: rounding and integer division in an index, a bound, a
// divisor and a string position (issue #286). Measured in Excel 16.0 64-bit
// (2026-10-02) through pyVBAharness.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function found(body: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    ${body}\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a value made whole before it is used', () => {
	it('is an index out of range', () => {
		for (const index of ['CInt(3.5)', 'Int(-0.1)', '2.6', '-0.6']) {
			expect(found(`Dim a(2) As Long\n    Main = a(${index})`), index).toEqual(['array-subscript-out-of-bounds']);
		}
	});

	it('is a ReDim bound below zero', () => {
		expect(found('Dim d() As Long\n    ReDim d(-0.6)\n    Main = 1')).toHaveLength(1);
	});

	it('is a zero divisor', () => {
		for (const divisor of ['Round(0.5)', 'Int(0.9)', 'Fix(-0.9)']) {
			expect(found(`Main = 10 / ${divisor}`), divisor).toEqual(['division-by-zero']);
		}
		expect(found('Dim x As Long\n    x = 1 \\ 3\n    Main = 10 / x')).toEqual(['division-by-zero']);
	});

	it('is a start or a code out of range', () => {
		expect(found('Main = Mid$("abc", CInt(0.5))')).toEqual(['runtime-argument-value']);
		expect(found('Main = Chr$(CInt(255.5))')).toEqual(['runtime-argument-value']);
	});

	it('stays quiet where it lands in range', () => {
		for (const body of [
			'Dim a(2) As Long\n    Main = a(CInt(2.5))',
			'Dim a(2) As Long\n    Main = a(Round(2.5))',
			'Dim a(2) As Long\n    Main = a(2.5)',
			'Dim a(2) As Long\n    Main = a(-0.5)',
			'Main = 10 / CInt(1.5)',
			'Main = 10 / Int(-0.9)',
			'Main = Chr$(CInt(254.5))',
			'Dim d() As Long\n    ReDim d(2.5)\n    Main = 1',
		]) {
			expect(found(body), body).toEqual([]);
		}
	});
});
