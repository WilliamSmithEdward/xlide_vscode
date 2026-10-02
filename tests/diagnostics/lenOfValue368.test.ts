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
