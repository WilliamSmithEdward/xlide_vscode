// Diagnostics tests: zero divisors that are not numbers (issue #491). Each
// raising sample was measured through pyVBAharness on 2026-10-02 in Excel
// 16.0 (build 20430); each quiet one runs there.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode } from '../helpers/diagnostics';

const ZERO = 'division-by-zero';

function wrap(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

describe('zero divisors that are not numbers (issue #491)', () => {
	it('reports a Boolean, a Date, Empty, a comparison and a String of "0"', () => {
		const bodies = [
			['Dim o As Boolean', 'Main = 5 / o'],
			['Dim o As Boolean', 'Main = 5 \\ o'],
			['Dim o As Boolean', 'Main = 5 Mod o'],
			['Dim o As Boolean', 'Main = 5 / CByte(o)'],
			['Dim o As Boolean', 'o = False', 'Main = 5 / o'],
			['Dim v As Variant', 'v = Empty', 'Main = 5 / v'],
			['Dim v As Variant', 'v = CDec(0)', 'Main = 5 / v'],
			['Dim v As Variant', 'Main = 5 / CLng(v)'],
			['Dim n As Long', 'n = 1', 'Main = 5 / (n = 2)'],
			['Main = 5 / (1 = 2)'],
			['Dim t As Date', 'Main = 5 / t'],
			['Dim s As String', 's = "0"', 'Main = 5 / s'],
			['Main = 5 / Empty'],
		];
		for (const body of bodies) {
			expect(byCode(analyzeModule(wrap(...body)), ZERO), body.join(' / ')).toHaveLength(1);
		}
	});

	it('stays quiet where the divisor is not zero', () => {
		const bodies = [
			['Dim o As Boolean', 'o = True', 'Main = 5 / o'],
			['Dim n As Long', 'n = 2', 'Main = 5 / (n = 2)'],
			['Dim s As String', 's = "2"', 'Main = 5 / s'],
			['Dim t As Date', 't = #1/1/2000#', 'Main = 5 / t'],
			['Dim v As Variant', 'v = True', 'Main = 5 / v'],
		];
		for (const body of bodies) {
			expect(byCode(analyzeModule(wrap(...body)), ZERO), body.join(' / ')).toHaveLength(0);
		}
	});

	it('reads a Boolean as -1 where an index or a length takes it', () => {
		expect(analyzeModule(wrap('Dim o As Boolean, a(-1 To 1) As Long', 'o = True', 'Main = a(o)')).filter((d) => d.severity === 'error')).toEqual([]);
		expect(analyzeModule(wrap('Dim o As Boolean', 'Main = Mid("abc", 1, o) & "x"')).filter((d) => d.severity === 'error')).toEqual([]);
	});
});
