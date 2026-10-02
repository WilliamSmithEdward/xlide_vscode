// Diagnostics tests: calls whose String result is fixed, or is a word no
// locale reads as a number, a Boolean or a date, stored into a typed local
// (issue #457). Measured in Excel 16.0 64-bit (build 20430, 2026-10-02)
// through pyVBAharness, each into a Long, a Double, a Boolean and a Date.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const TARGETS = ['Long', 'Double', 'Boolean', 'Date'];

function found(expr: string, target: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    Dim x As ${target}\n    x = ${expr}\n    Main = x\nEnd Function\n`;
	return analyzeModule(src, { host: 'excel' }).filter((diag) => diag.severity === 'error').map((diag) => `${diag.code}: ${diag.message}`);
}

describe('a String result that does not convert', () => {
	it('raises 13 into a Long, a Double, a Boolean and a Date', () => {
		for (const expr of [
			'Chr(65)', 'Chr$(66)', 'Hex(255)', 'Space(2)', 'Space(0)', 'String(3, "x")', 'StrConv("abc", vbProperCase)',
			'StrConv("ab", vbUpperCase)', 'MonthName(1)', 'MonthName(2, True)', 'WeekdayName(1)', 'Format(#1/2/2000#, "mmm")',
			'TypeName(Range("A1"))', 'TypeName(1)', 'Range("A1").Address', 'Range("A1:B2").Address(False, False)',
			'Cells(1, 2).Address', 'Application.Name', 'Application.PathSeparator',
		]) {
			for (const target of TARGETS) {
				expect(found(expr, target), `${expr} into ${target}`).toEqual([expect.stringMatching(/^assignment-type-mismatch: .*'13'/)]);
			}
		}
	});

	it('names what the locale or host decides', () => {
		expect(found('MonthName(1)', 'Long')).toEqual([expect.stringContaining('holds a month name')]);
		expect(found('Range("A1").Address', 'Long')).toEqual([expect.stringContaining('holds a cell address')]);
		expect(found('Chr(65)', 'Long')).toEqual([expect.stringContaining('holds "A"')]);
	});
});

describe('a String result that converts', () => {
	it('is quiet', () => {
		for (const expr of ['Hex(9)', 'Format(12, "0")', 'Application.Version', 'Range("A5").Row', 'Chr(55)', 'Oct(7)', 'Oct(64)', 'String(2, "7")']) {
			for (const target of TARGETS) {
				expect(found(expr, target), `${expr} into ${target}`).toEqual([]);
			}
		}
	});
});
