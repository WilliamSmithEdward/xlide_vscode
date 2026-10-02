// Diagnostics tests: Len and LenB of a value of known type other than String
// or Variant need a variable (issue #368). Measured in Excel 16.0 (build
// 20326, 2026-10-01) with Debug > Compile.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

const CODE = 'variable-required';
const HELPERS = 'Private Function FL() As Long\n    FL = 5\nEnd Function\nPrivate Function FV() As Variant\n    FV = 5\nEnd Function\nPrivate Function FS() As String\n    FS = "ab"\nEnd Function\n';

function source(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n${HELPERS}`;
}

describe('Len of a value that is not a String', () => {
	it('is Variable required', () => {
		for (const fn of ['Len', 'LenB']) {
			for (const [lines, span] of [
				[[`Main = ${fn}(1.5)`], '1.5'],
				[[`Main = ${fn}(40000)`], '40000'],
				[['Dim i As Long', `Main = ${fn}(i + 1)`], 'i + 1'],
				[['Dim i As Long', `Main = ${fn}(i = 1)`], 'i = 1'],
				[[`Main = ${fn}(Val("3"))`], 'Val("3")'],
				[[`Main = ${fn}(#1/2/2000#)`], '#1/2/2000#'],
			] as const) {
				const src = source(...lines);
				expectDiagnostic(src, byCode(analyzeModule(src), CODE), CODE, { span, message: 'Variable required' });
			}
		}
		for (const [lines, span] of [
			[['Const K = 5', 'Main = Len(K)'], 'K'],
			[['Main = Len(FL())'], 'FL()'],
			[['Main = Len(CLng(3))'], 'CLng(3)'],
			[['Main = Len(True)'], 'True'],
		] as const) {
			const src = source(...lines);
			expectDiagnostic(src, byCode(analyzeModule(src), CODE), CODE, { span });
		}
	});

	it('compiles for a String, a Variant, a variable or an element', () => {
		for (const lines of [
			['Main = Len("abc")'],
			['Dim i As Long', 'Main = Len("a" & i)'],
			['Main = Len(CStr(12))'],
			['Main = Len(Left("abc", 2))'],
			['Main = Len(Mid("abc", 2))'],
			['Dim d As Double', 'Main = Len(d)'],
			['Dim d As Double', 'Main = Len((d))'],
			['Dim v As Variant', 'v = 12', 'Main = Len(v)'],
			['Dim a(2) As Long', 'Main = Len(a(1))'],
			['Main = Len(Now)'],
			['Main = Len(FV())'],
			['Main = Len(FS())'],
		]) {
			expect(byCode(analyzeModule(source(...lines)), CODE), lines.join(' : ')).toEqual([]);
		}
	});
});

// Measured in Excel 16.0 64-bit (build 20430, 2026-10-02) with Debug > Compile (issue #455).
describe('Len of an expression with a Variant in it', () => {
	const DIMS = ['Dim v As Variant, c As Currency, i As Integer, d As Double, s As String', 'Dim g As Single, b As Byte, l As Long'];

	it('compiles when an operand or the function result is a Variant', () => {
		for (const expr of [
			'c < v', 'i \\ Round(1.5)', 'v Like "a"', 'i \\ v', 'i Mod Round(1.5)', 'i + Round(1.5)', 'i And Round(1.5)',
			'Not Round(1.5)', 'Fix(i)', 'Int(l)', 'Fix(v)', 'Abs(b)', 'Abs(Round(1.5))', 'v & 1',
		]) {
			expect(byCode(analyzeModule(source(...DIMS, `Main = Len(${expr})`)), CODE), expr).toEqual([]);
		}
	});

	it('is Variable required for a typed result, named by its type', () => {
		for (const [expr, type] of [
			['i \\ 2', 'an Integer'], ['i \\ d', 'a Long'], ['d Mod 2', 'a Long'], ['i Like s', 'a Boolean'], ['s < "a"', 'a Boolean'],
			['Fix(d)', 'a Double'], ['Int(g)', 'a Single'], ['Fix(c)', 'a Currency'], ['Fix(2)', 'an Integer'], ['Abs(l)', 'a Long'],
			['Abs(s)', 'a Double'], ['Sgn(v)', 'an Integer'], ['Sqr(v)', 'a Double'], ['-d', 'a Double'], ['Not i', 'an Integer'],
		] as const) {
			const src = source(...DIMS, `Main = Len(${expr})`);
			expectDiagnostic(src, byCode(analyzeModule(src), CODE), CODE, { span: expr, message: `Len of ${type} ` });
		}
	});
});
