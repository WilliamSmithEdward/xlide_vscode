// Diagnostics tests: a String added to a Date or Boolean, and `^` with a
// local's value (issue #331). Measured in Excel 16.0 64-bit (build 20430,
// 2026-10-02) through pyVBAharness.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(...lines: string[]): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    Dim s As String, d As Date, b As Boolean, z As Long, x As Double\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => `${diag.code}: ${diag.message}`);
}

describe('operator leftovers', () => {
	it('raise 13 for a String added to a Date or Boolean that cannot read it', () => {
		for (const lines of [['s = "x"', 'd = #1/1/2000#', 'Main = s + d'], ['s = "x"', 'b = True', 'Main = s + b'], ['s = ""', 'd = #1/1/2000#', 'Main = s + d']]) {
			expect(errors(...lines), lines.join(' / ')).toEqual([expect.stringMatching(/^string-arithmetic-coercion: .*'13'/)]);
		}
		expect(errors('s = "1/1/2000"', 'd = #1/1/2000#', 'Main = s + d')).toEqual([]);
	});

	it('raise 5 for ^ on a local known to hold a value that refuses it', () => {
		for (const lines of [['Main = z ^ -1'], ['x = -4', 'Main = x ^ 0.5'], ['Main = 0 ^ True']]) {
			expect(errors(...lines), lines.join(' / ')).toEqual([expect.stringContaining("'5'")]);
		}
		for (const lines of [['x = 4', 'Main = x ^ 0.5'], ['x = -8', 'Main = x ^ 2']]) {
			expect(errors(...lines), lines.join(' / ')).toEqual([]);
		}
	});
});
