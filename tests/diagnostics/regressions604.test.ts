// Diagnostics tests: #604, DateSerial carrying the month into the year
// first, and a String into a Type's dynamic array field. Measured on
// 2026-10-02 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const TYPE = 'Private Type T\n    f() As Long\n    b() As Byte\nEnd Type\n';

function raised(body: string, head = ''): string[] {
	const source = `Option Explicit\n${head}Function Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(source)
		.filter((diag) => diag.severity === 'error')
		.map((diag) => /Run-time error '(-?\d+)'/.exec(diag.message)?.[1] ?? diag.code);
}

describe('DateSerial carries the month into the year, then reads the year (issue #604)', () => {
	it('runs where the carried year lands in range', () => {
		for (const args of ['100, 0, 1', '100, -1200, 1', '101, -13, 1', '366, -10000, -1', '10000, 0, 1', '9999, 12, 31', '100, 1, 1', '-1900, 1, 1', '0, 1, 1', '29, 12, 31', '99, 12, 31', '1, -11, 1']) {
			expect(raised(`Main = DateSerial(${args})`), args).toEqual([]);
		}
	});

	it('raises 5 past either end, after the carry', () => {
		for (const args of ['9999, 13, 0', '10000, 1, 0', '99, 13, -1', '9999, 12, 32', '100, 1, 0', '-1901, 12, 31', '-10000, 1, 1', '9999, 13, 1']) {
			expect(raised(`Main = DateSerial(${args})`), args).toEqual(['5']);
		}
	});
});

describe('a Type\'s dynamic array field given a value (issue #604)', () => {
	it('refuses a String unless the field is As Byte', () => {
		for (const body of ['Dim t As T\n    t.f = "abcdef"', 'Dim ts(1) As T\n    ts(0).f = "abcdef"', 'Dim t As T\n    With t\n        .f = "abcdef"\n    End With']) {
			expect(raised(body, TYPE), body).toEqual(['array-target-assignment']);
		}
		expect(raised('Dim t As T\n    t.b = "abcdef"\n    Main = UBound(t.b)', TYPE)).toEqual([]);
	});

	it('raises 13 for Split into a Long field', () => {
		expect(raised('Dim t As T\n    t.f = Split("a b")', TYPE)).toEqual(['13']);
	});
});
