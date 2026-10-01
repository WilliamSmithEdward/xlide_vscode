// Diagnostics tests: a String expression the analyzer can spell out, assigned
// to a Boolean, a number or a Date, is checked like a literal (issue #405).
// Every case was measured in Excel 16.0 (build 20326, 2026-10-01).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

const CODE = 'assignment-type-mismatch';

function source(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${[...lines, 'Main = 1'].map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

describe('a spelled-out String expression into a typed local', () => {
	it('raises 13 where the text does not convert', () => {
		for (const [lines, span] of [
			[['Dim r As Boolean', 'r = "a" & "b"'], '"a" & "b"'],
			[['Dim r As Boolean', 'r = Left("abc", 1)'], 'Left("abc", 1)'],
			[['Dim r As Boolean', 'r = CStr(#1/2/2000#)'], 'CStr(#1/2/2000#)'],
			[['Dim r As Boolean', 'Dim t As Date', 't = #1/2/2000#', 'r = t & 5'], 't & 5'],
			[['Dim r As Boolean', 'Dim t As Date', 'r = t & ""'], 't & ""'],
			[['Dim r As Boolean', 'r = #1/2/2000# & ""'], '#1/2/2000# & ""'],
			[['Dim r As Boolean', 'Dim s As String', 's = "a"', 'r = s & "b"'], 's & "b"'],
			[['Dim r As Boolean', 'r = Mid("ab12", 1, 2)'], 'Mid("ab12", 1, 2)'],
			[['Dim r As Boolean', 'r = " True "'], '" True "'],
			[['Dim r As Boolean', 'r = LTrim(" True ")'], 'LTrim(" True ")'],
			[['Dim n As Long', 'n = "a" & "b"'], '"a" & "b"'],
			[['Dim n As Long', 'n = CStr(#1/2/2000#)'], 'CStr(#1/2/2000#)'],
			[['Dim d As Date', 'd = "a" & "b"'], '"a" & "b"'],
		] as const) {
			const src = source(...lines);
			expectDiagnostic(src, byCode(analyzeModule(src), CODE), CODE, { span, message: "'13'" });
		}
	});

	it('runs where the text converts, or is not known', () => {
		for (const lines of [
			['Dim r As Boolean', 'r = 1 & 2'],
			['Dim r As Boolean', 'r = "Tr" & "ue"'],
			['Dim r As Boolean', 'r = Format(1, "0.0")'],
			['Dim r As Boolean', 'r = Left("Trueish", 4)'],
			['Dim r As Boolean', 'r = UCase("true")'],
			['Dim r As Boolean', 'r = LCase("FALSE")'],
			['Dim r As Boolean', 'r = Right("xTrue", 4)'],
			['Dim r As Boolean', 'r = Trim("  5 ")'],
			['Dim r As Boolean', 'r = Trim(" True ")'],
			['Dim r As Boolean', 'r = RTrim(LTrim(" True "))'],
			['Dim r As Boolean', 'r = Mid("xx12", 3)'],
			['Dim r As Boolean', 'r = True & ""'],
			['Dim r As Boolean', 'Dim s As String', 's = "Tr"', 'r = s & "ue"'],
			['Dim n As Long', 'n = 1 & 2'],
			['Dim n As Long', 'n = Left("12x", 2)'],
			['Dim n As Integer', 'n = 4 & 0000'],
			['Dim n As Double', 'n = "1" & "." & "5"'],
			['Dim d As Date', 'd = CStr(#1/2/2000#)'],
		]) {
			expect(byCode(analyzeModule(source(...lines)), CODE), lines.join(' : ')).toEqual([]);
		}
	});
});
