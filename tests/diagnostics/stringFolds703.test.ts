// Diagnostics tests: string conversions the rules fold where every locale
// reads the string alike (issue #703). Measured on 2026-10-03 in Excel 16.0
// (build 20430), where the comma groups thousands.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { numericStringReadings, valPrefixValue } from '../../src/analyzer/diagnostics/stringConversion';

function errors(body: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    ${body}\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('Val, as VBA reads a string', () => {
	it.each([
		['0,5', 0], ['abc', 0], ['1,000', 1], ['1 2 3', 123], ['&H10', 16], [' -5.5x', -5.5], ['1e3', 1000], ['.5', 0.5], ['', 0], ['-', 0],
		['1.2.3', 1.2], ['1d2', 100], ['- 5', -5], ['12abc', 12],
	])('Val(%j) is %d', (text, value) => {
		expect(valPrefixValue(text)).toBe(value);
	});
});

describe('a string under each decimal point', () => {
	it('reads both ways, or not at all', () => {
		expect(numericStringReadings('3.5')).toEqual([3.5, 35]);
		expect(numericStringReadings('1,5')).toEqual([15, 1.5]);
		expect(numericStringReadings('4')).toEqual([4, 4]);
		expect(numericStringReadings('1.2.3')).toBeUndefined();
		expect(numericStringReadings('$5')).toBeUndefined();
	});
});

describe('string conversions every locale reads alike (issue #703)', () => {
	it.each([
		['10 \\ CLng("&H0")', 'Main = 10 \\ CLng("&H0")', 'division-by-zero'],
		['10 / CLng("&H0")', 'Main = 10 / CLng("&H0")', 'division-by-zero'],
		['10 Mod CInt("&O0")', 'Main = 10 Mod CInt("&O0")', 'division-by-zero'],
		['10 \\ Val("0,5")', 'Main = 10 \\ Val("0,5")', 'division-by-zero'],
		['10 \\ Val("0.4")', 'Main = 10 \\ Val("0.4")', 'division-by-zero'],
		['10 \\ Val(" 0 ")', 'Main = 10 \\ Val(" 0 ")', 'division-by-zero'],
		['10 \\ Val("abc")', 'Main = 10 \\ Val("abc")', 'division-by-zero'],
		['CByte("(5)")', 'Main = CByte("(5)")', 'arithmetic-overflow'],
		['CByte("5-")', 'Main = CByte("5-")', 'arithmetic-overflow'],
		['CInt("(40000)")', 'Main = CInt("(40000)")', 'arithmetic-overflow'],
		['a(CInt("4"))', 'Dim a(3) As Long\n    Main = a(CInt("4"))', 'array-subscript-out-of-bounds'],
		['a(CLng("&H4"))', 'Dim a(3) As Long\n    Main = a(CLng("&H4"))', 'array-subscript-out-of-bounds'],
		['a(CInt("3.5")), 4 or 35', 'Dim a(3) As Long\n    Main = a(CInt("3.5"))', 'array-subscript-out-of-bounds'],
	])('reports %s', (_label, body, code) => {
		expect(errors(body), body).toEqual([code]);
	});

	it.each([
		// The locale decides these: 40000 or 40, 0 or 4, and a month or a day.
		'Main = CInt("40,000")',
		'Main = 10 \\ CInt("0.4")',
		'Main = CDate("2000-13-01")',
		'Main = 10 \\ CDbl("0,5")',
		'Dim a(3) As Long\n    Main = a(CInt("2.5"))',
		'Main = CInt("3,000")',
		'Main = CInt("&H8000")',
		'Main = Val("&H10")',
		'Main = Val("1,000")',
		'Main = Val("1 2 3")',
		'Main = CCur("1,234.5")',
		'Main = CInt("2.5")',
		'Main = CByte("(0)")',
	])('stays quiet on %s', (body) => {
		expect(errors(body), body).toEqual([]);
	});
});
