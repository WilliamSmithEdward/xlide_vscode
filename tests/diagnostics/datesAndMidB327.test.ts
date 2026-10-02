// Diagnostics tests: Date, Now and DateSerial into an Integer, and the MidB
// statement past the end (issue #327). Measured in Excel 16.0 64-bit (build
// 20430, 2026-10-02) through pyVBAharness.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(...lines: string[]): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    Dim i As Integer, n As Long, s As String\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => `${diag.code}: ${diag.message}`);
}

describe('a date into an Integer', () => {
	it('overflows from Date, Now and DateSerial', () => {
		for (const line of ['i = Date', 'i = Now', 'Main = CInt(Now)', 'i = DateSerial(2020, 1, 1)']) {
			expect(errors(line), line).toEqual([expect.stringContaining("'6'")]);
		}
	});

	it('runs into a Long, or from TimeSerial', () => {
		for (const line of ['n = Now', 'n = Date', 'i = TimeSerial(12, 0, 0)', 'i = Date - DateSerial(2020, 1, 1)']) {
			expect(errors(line), line).toEqual([]);
		}
	});
});

describe('the MidB statement', () => {
	it('raises 5 past the string\'s bytes', () => {
		expect(errors('s = "abc"', 'MidB(s, 9, 1) = "x"')).toEqual([expect.stringMatching(/MidB statement start 9 .* 6 byte\(s\) long.*'5'/)]);
		expect(errors('s = "abc"', 'MidB(s, 5, 1) = "x"')).toEqual([]);
	});
});
