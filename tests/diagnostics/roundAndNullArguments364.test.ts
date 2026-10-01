// Diagnostics tests: Round's digit limit, which follows the value's type,
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

// The limit follows the first argument's type, not its fraction (issue
// #402, measured in Excel 16.0 64-bit, build 20430, 2026-10-01).
describe("Round's NumDigitsAfterDecimal", () => {
	it('runs past 22 for an Integer, Long, Boolean, Currency, Decimal or Null', () => {
		for (const lines of [
			['Main = Round(1, 23)'],
			['Main = Round(256, 40000)'],
			['Main = Round(0, 30)'],
			['Main = Round(3, 256)'],
			['Main = Round(10 \\ 2, 23)'],
			['Main = Round(CCur(3.5), 256)'],
			['Main = Round(CDec(3.5), 256)'],
			['Main = Round(CLng(3.5), 23)'],
			['Main = Round(True, 23)'],
			['Dim n As Long', 'n = 3', 'Main = Round(n, 23)'],
			['Dim v As Variant', 'v = 3', 'Main = Round(v, 23)'],
			['Main = IsNull(Round(Null, 30))'],
			['Main = Round(1.5, 22)'],
		]) {
			expect(byCode(analyzeModule(source(...lines)), CODE), lines.join(' : ')).toEqual([]);
		}
	});

	it('raises past 22 for a Double, Single, String or Date, and below 0 for any', () => {
		for (const lines of [
			['Main = Round(1.5, 23)'],
			['Main = Round(-2.5, 23)'],
			['Main = Round(3#, 23)'],
			['Main = Round(CDbl(3), 23)'],
			['Main = Round(0#, 23)'],
			['Main = Round(1E2, 23)'],
			['Main = Round(10 / 2, 23)'],
			['Main = Round(CSng(3), 23)'],
			['Main = Round("3", 23)'],
			['Main = Round(#1/2/2000#, 23)'],
			['Dim d As Double', 'd = 3', 'Main = Round(d, 23)'],
			['Dim d As Double', 'Main = Round(d, 23)'],
			['Dim v As Variant', 'v = 3#', 'Main = Round(v, 23)'],
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
