// Diagnostics tests: Round's digit limit binds only a value with a fraction,
// and a Null string argument returns Null before the others are checked
// (issue #364). Every case was measured in Excel 16.0 (build 20326,
// 2026-10-01).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

const CODE = 'runtime-argument-value';

function source(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

describe("Round's NumDigitsAfterDecimal", () => {
	it('runs past 22 for a whole number or Null', () => {
		for (const line of ['Main = Round(1, 23)', 'Main = Round(256, 40000)', 'Main = Round(0, 30)', 'Main = IsNull(Round(Null, 30))', 'Main = Round(1.5, 22)']) {
			expect(byCode(analyzeModule(source(line)), CODE), line).toEqual([]);
		}
	});

	it('raises past 22 for a value with a fraction, and below 0 for any', () => {
		for (const lines of [
			['Main = Round(1.5, 23)'],
			['Main = Round(-2.5, 23)'],
			['Dim x As Double', 'x = 1.5', 'Main = Round(x, 23)'],
			['Main = Round(1, -1)'],
		]) {
			const src = source(...lines);
			expectDiagnostic(src, byCode(analyzeModule(src), CODE), CODE, { message: "'5'" });
		}
	});
});

describe('a Null string argument', () => {
	it('returns Null before the other arguments are checked', () => {
		for (const lines of [
			['Main = IsNull(Mid(Null, 0))'],
			['Main = IsNull(Mid(Null, 1, -1))'],
			['Main = IsNull(Left(Null, -1))'],
			['Main = IsNull(Right(Null, -1))'],
			['Main = IsNull(InStr(0, Null, "a"))'],
			['Main = IsNull(InStr(0, "a", Null))'],
			['Main = IsNull(StrComp(Null, "a", -9))'],
			['Dim v As Variant', 'v = Null', 'Main = IsNull(Mid(v, 0))'],
		]) {
			expect(byCode(analyzeModule(source(...lines)), CODE), lines.join(': ')).toEqual([]);
		}
	});

	it('leaves a String argument checked', () => {
		for (const line of ['Main = Mid("abc", 0)', 'Main = Left("abc", -1)']) {
			const src = source(line);
			expectDiagnostic(src, byCode(analyzeModule(src), CODE), CODE, { message: "'5'" });
		}
	});
});
