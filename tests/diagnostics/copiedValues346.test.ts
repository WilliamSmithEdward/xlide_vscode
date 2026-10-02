// Diagnostics tests: a value known from a literal is followed through a
// copy into another local (issue #346). Measured in Excel 16.0 64-bit (build
// 20430, 2026-10-02) through pyVBAharness.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(...lines: string[]): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => `${diag.code}: ${diag.message}`);
}

describe('a value copied into another local', () => {
	it('is followed into the rules that read it', () => {
		for (const [lines, error] of [
			[['Dim a As Long, d As Long', 'a = 0', 'd = a', 'Main = 10 / d'], "'11'"],
			[['Dim a(3) As Long, i As Long, j As Long', 'i = 5', 'j = i', 'Main = a(j)'], "'9'"],
			[['Dim a As String, s As String', 'a = "abc"', 's = a', 'Main = s + 1'], "'13'"],
			[['Dim a As Double, d As Double', 'a = 0.4', 'd = a', 'Main = 10 Mod d'], "'11'"],
		] as const) {
			expect(errors(...lines), lines.join(' / ')).toEqual([expect.stringContaining(error)]);
		}
	});

	it('reads a String local known to hold a key', () => {
		expect(errors('Dim c As New Collection, k As String', 'c.Add 1, "a"', 'k = "zz"', 'Main = c(k)')).toEqual([expect.stringMatching(/^collection-key-not-found: .*'5'/)]);
		expect(errors('Dim c As New Collection, k As String', 'c.Add 1, "a"', 'k = "A"', 'Main = c(k)')).toEqual([]);
	});

	it('keeps the value the copy took, whatever the source holds later', () => {
		expect(errors('Dim a As Long, d As Long', 'a = 5', 'd = a', 'a = 0', 'Main = 10 / d')).toEqual([]);
		expect(errors('Dim a As Long, d As Long', 'a = 0', 'd = a', 'd = 2', 'Main = 10 / d')).toEqual([]);
	});
});
