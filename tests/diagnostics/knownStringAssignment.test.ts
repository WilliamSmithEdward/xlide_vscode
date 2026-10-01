// Diagnostics tests: a String the analyzer knows a local or an array element
// holds, assigned to a number, Boolean or Date local. Every case was measured
// in Excel 16.0 (build 20326, 2026-10-01).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

const CODE = 'assignment-type-mismatch';

function source(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\n    Main = n\nEnd Function\n`;
}

describe('a known String assigned to a scalar local', () => {
	it('reports a value the target cannot take', () => {
		for (const [lines, span, message] of [
			[['Dim s As String, n As Long', 's = "b"', 'n = s'], 's', "'13'"],
			[['Dim v As Variant, n As Long', 'v = "b"', 'n = v'], 'v', "'13'"],
			[['Dim v, n As Long', 'v = "b"', 'n = v'], 'v', "'13'"],
			[['Dim v As Variant, n As Long', 'v = Array("1", "b")', 'n = v(1)'], 'v(1)', "'13'"],
			[['Dim v As Variant, n As Double', 'v = Split("1,b", ",")', 'n = v(1)'], 'v(1)', "'13'"],
			[['Dim s As String, n As Byte', 's = "b"', 'n = (s)'], 's', "'13'"],
			[['Dim s As String, n As Long', 'n = s'], 's', 'holds ""'],
			[['Dim s As String, n As Integer', 's = "True"', 'n = s'], 's', "'13'"],
			[['Dim s As String, n As Integer', 's = "&H10000"', 'n = s'], 's', "'6'"],
			[['Dim s As String, n As Boolean', 's = "yes"', 'n = s'], 's', "'13'"],
			[['Dim s As String, n As Boolean', 's = ""', 'n = s'], 's', "'13'"],
			[['Dim v As Variant, n As Date', 'v = Array("1", "b")', 'n = v(1)'], 'v(1)', "'13'"],
			[['Dim s As String, n As Date', 's = "True"', 'n = s'], 's', "'13'"],
		] as const) {
			const src = source(...lines);
			expectDiagnostic(src, byCode(analyzeModule(src), CODE), CODE, { span, message });
		}
	});

	it('stays quiet when the value converts or the target takes any String', () => {
		for (const lines of [
			['Dim s As String, n As Long', 's = "5"', 'n = s'],
			['Dim s As String, n As Long', 's = " 2 "', 'n = s'],
			['Dim s As String, n As Long', 's = "&H10"', 'n = s'],
			['Dim s As String, n As Long', 's = "1e3"', 'n = s'],
			['Dim s As String, n As Long', 's = "$5"', 'n = s'],
			['Dim s As String, n As Long', 's = "&H10000"', 'n = s'],
			['Dim v As Variant, n As Long', 'v = Array("1", "b")', 'n = v(0)'],
			['Dim s As String, n As Long', 's = "b"', 's = "5"', 'n = s'],
			['Dim s As String, n As Long', 's = "b"', 'Mid(s, 1, 1) = "5"', 'n = s'],
			['Dim v As Variant, n As Long', 'n = v'],
			['Dim s As String, n As String', 's = "b"', 'n = s'],
			['Dim s As String, n As Variant', 's = "b"', 'n = s'],
			['Dim s As String, n As Date', 's = "1/2/2020"', 'n = s'],
			['Dim s As String, n As Date', 's = "5"', 'n = s'],
			// Not the element alone, an element written after the Array, and
			// an assignment a known guard keeps from running.
			['Dim v As Variant, n As Long', 'v = Array("1", "")', 'n = v(1) & "5"'],
			['Dim v As Variant, n As Long', 'v = Array("1", "b")', 'v(1) = "5"', 'n = v(1)'],
			['Dim s As String, n As Long, k As Long', 's = "b"', 'If k > 0 Then n = s'],
			['Dim s As String, n As Long, k As Long', 's = "b"', 'If k > 0 Then', '    n = s', 'End If'],
		]) {
			expect(byCode(analyzeModule(source(...lines)), CODE), lines.join(': ')).toEqual([]);
		}
	});

	it('does not report a String of unknown value into a Boolean', () => {
		for (const value of ['"True"', '"5"', '" 2 "', '"&H10"', '"$5"']) {
			const lines = ['Dim s As String, n As Boolean', `s = ${value}`, 'n = s'];
			expect(byCode(analyzeModule(source(...lines)), CODE), value).toEqual([]);
		}
		const unknown = 'Option Explicit\nFunction Main(s As String) As Boolean\n    Dim n As Boolean\n    n = s\n    Main = n\nEnd Function\n';
		expect(byCode(analyzeModule(unknown), CODE)).toEqual([]);
	});
});
