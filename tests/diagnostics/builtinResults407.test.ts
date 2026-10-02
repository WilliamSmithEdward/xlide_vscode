// Diagnostics tests: a built-in's result stored into a narrower type, and a
// String local into a conversion (issue #407). Measured in Excel 16.0
// 64-bit (build 20430, 2026-10-02) through pyVBAharness.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(...lines: string[]): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    Dim r As Byte, l As Long, s As String\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => `${diag.code}: ${diag.message}`);
}

describe('a built-in result that does not fit', () => {
	it('overflows', () => {
		for (const lines of [
			['r = Sgn(-5)'], ['r = Choose(1, 300)'], ['r = IIf(True, 300, 0)'], ['r = Len(String(300, "a"))'], ['r = AscW(ChrW(300))'],
			['l = -5', 'r = Sgn(l)'], ['s = "40000"', 'Main = CInt(s)'], ['s = "1E3"', 'Main = CByte(s)'],
		]) {
			expect(errors(...lines), lines.join(' / ')).toEqual([expect.stringContaining("'6'")]);
		}
	});

	// Issue #351: Asc, AscW and Sgn return an Integer, so Integer arithmetic on
	// them overflows.
	it('overflows Integer arithmetic on Asc, AscW and Sgn', () => {
		for (const line of ['Main = Asc("a") * 1000', 'Main = AscW("a") * 1000', 'Main = Sgn(1) * 32767 + 1']) {
			expect(errors(line), line).toEqual([expect.stringContaining("'6'")]);
		}
		for (const line of ['Main = Asc("a") * 300', 'Main = Year(#12/1/2026#) * 100 + Month(#12/1/2026#)', 'Main = Round(CInt(300)) * 300']) {
			expect(errors(line), line).toEqual([]);
		}
	});

	it('is quiet where it fits', () => {
		for (const lines of [['r = Sgn(5)'], ['r = Choose(2, 300, 3)'], ['r = IIf(False, 300, 0)'], ['r = Len(String(30, "a"))'], ['s = "300"', 'Main = CInt(s)']]) {
			expect(errors(...lines), lines.join(' / ')).toEqual([]);
		}
	});
});
