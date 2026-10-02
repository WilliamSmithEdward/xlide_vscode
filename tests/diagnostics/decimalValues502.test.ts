// Diagnostics tests: Decimal values followed past CDec (issue #502). Each
// sample was measured through pyVBAharness on 2026-10-02 in Excel 16.0
// (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(...lines: string[]): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
	return analyzeModule(src).filter((d) => d.severity === 'error').map((d) => d.code ?? '');
}

const MAX = 'CDec("79228162514264337593543950335")';

describe('Decimal values past CDec (issue #502)', () => {
	it('overflows past the largest Decimal, and where a Long or Currency takes one', () => {
		const bodies = [
			[`Main = ${MAX} + 1`],
			[`Main = ${MAX} * 2`],
			['Dim m As Variant', `m = ${MAX}`, 'Main = m + 1'],
			['Dim m As Variant, n As Long', `m = ${MAX}`, 'n = m'],
			['Dim m As Variant, c As Currency', `m = ${MAX}`, 'c = m'],
			['Dim m As Variant', `m = ${MAX}`, 'Main = CLng(m)'],
			['Dim m As Variant', `m = ${MAX}`, 'Main = m And 1'],
		];
		for (const body of bodies) {
			expect(errors(...body), body.join(' / ')).toEqual(['arithmetic-overflow']);
		}
	});

	it('names the overflow, not a division by zero, where Mod converts first', () => {
		expect(errors('Dim m As Variant', `m = ${MAX}`, 'Main = m Mod 0')).toEqual(['arithmetic-overflow']);
		expect(errors('Main = 1E10 Mod 0')).toEqual(['arithmetic-overflow']);
		expect(errors('Dim d As Long', 'Main = 5 Mod d')).toEqual(['division-by-zero']);
	});

	it('stays quiet on Decimals that fit', () => {
		for (const body of [[`Main = ${MAX}`], ['Main = CDec(1.5) * 2'], ['Main = CInt(CDec(3))'], ['Dim m As Variant', `m = ${MAX}`, 'Main = m - 1'], ['Main = CDec("79228162514264337593543950334") + 1']]) {
			expect(errors(...body), body.join(' / ')).toEqual([]);
		}
	});
});
