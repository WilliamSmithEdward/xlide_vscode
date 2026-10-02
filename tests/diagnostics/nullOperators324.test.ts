// Diagnostics tests: Null carried through an operator into a typed target
// raises 94 (issue #324). Measured in Excel 16.0 64-bit (build 20430,
// 2026-10-02) through pyVBAharness.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(...lines: string[]): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    Dim s As String, n As Long, b As Boolean, v As Variant, w As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => `${diag.code}: ${diag.message}`);
}

describe('Null through an operator into a typed target', () => {
	it('raises 94', () => {
		for (const lines of [
			['s = "a" + Null'], ['s = Null + "a"'], ['n = 1 + Null'], ['n = Null * 2'], ['n = -Null'], ['n = Abs(Null)'],
			['b = Not Null'], ['b = True And Null'], ['b = (1 = Null)'], ['v = Null', 'n = v + 1'], ['v = Null', 's = "a" + v'],
		]) {
			expect(errors(...lines), lines.join(' / ')).toEqual([expect.stringMatching(/^assignment-type-mismatch: .*'94'/)]);
		}
	});

	it('is quiet where the result is not Null, or the target is a Variant', () => {
		for (const lines of [
			['s = "a" & Null'], ['s = Null & "a"'], ['b = False And Null'], ['b = True Or Null'], ['w = 1 + Null'],
			['v = "b"', 's = "a" + v'],
		]) {
			expect(errors(...lines), lines.join(' / ')).toEqual([]);
		}
	});
});
