// Diagnostics tests: Not binds below the comparisons, and a String of unknown
// value into a Boolean (issue #361). Every case was measured in Excel 16.0
// (build 20326, 2026-10-01).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

const COERCION = 'string-arithmetic-coercion';

function source(lines: readonly string[], extra = ''): string {
	return `Option Explicit\n${extra}Function Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

function errors(src: string): string[] {
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('Not with a comparison in its operand', () => {
	it('negates the comparison, a Boolean', () => {
		for (const lines of [
			['Dim s As String', 's = "x"', 'Main = Not s = "y"'],
			['Dim s As String', 's = "abc"', 'Main = Not s Like "a*"'],
			['Dim v As Variant', 'v = "x"', 'Main = Not v = 1'],
			['Dim v As Variant', 'v = "x"', 'Main = Not v < 2'],
			['Dim v As Variant', 'v = "x"', 'Main = Not v = 1 And True'],
			['Dim s As String', 's = "x"', 'Main = Not s = "y" Or s = "x"'],
			['Dim s As String', 's = "x"', 'If Not s = "Sheet1" Then Main = 1'],
		]) {
			expect(errors(source(lines)), lines.join(': ')).toEqual([]);
		}
	});

	it('still reports Not of the string itself', () => {
		for (const [lines, span] of [
			[['Dim s As String', 's = "x"', 'Main = Not s'], 's'],
			[['Dim s As String', 's = "x"', 'Main = (Not s) = "y"'], 's'],
			[['Dim s As String', 's = "x"', 'Main = Not s & "y"'], 's'],
			[['Dim v As Variant', 'v = "x"', 'Main = Not (v)'], 'v'],
			// A comparison inside a call is no comparison of Not's operand.
			[['Dim s As String', 's = "x"', 'Main = Not s & CStr(1 = 1)'], 's'],
		] as const) {
			const src = source(lines);
			expectDiagnostic(src, byCode(analyzeModule(src), COERCION), COERCION, { span, message: "'13'" });
		}
		const and = source(['Dim s As String', 's = "x"', 'Main = Not s And True']);
		expect(byCode(analyzeModule(and), COERCION).length).toBeGreaterThan(0);
	});
});

describe('a String of unknown value into a Boolean', () => {
	it('is not reported', () => {
		for (const [lines, extra] of [
			[['Dim b As Boolean', 'b = GetSetting("app", "s", "k", "False")', 'Main = b'], ''],
			[['Dim b As Boolean', 'b = Cfg("enabled")', 'Main = b'], 'Private Function Cfg(k As String) As String\n    Cfg = "True"\nEnd Function\n'],
			[['Dim b As Boolean', 'b = CStr(2)', 'Main = b'], ''],
			[['Dim b As Boolean', 'b = Format(1)', 'Main = b'], ''],
			[['Dim b As Boolean', 'b = Left("1x", 1)', 'Main = b'], ''],
			[['Dim b As Boolean, d As Double', 'd = 0', 'b = d & d', 'Main = b'], ''],
			[['Main = TakeB(CStr(1))'], 'Private Function TakeB(ByVal b As Boolean) As Boolean\n    TakeB = b\nEnd Function\n'],
		] as const) {
			expect(errors(source(lines, extra)), lines.join(': ')).toEqual([]);
		}
	});
});
