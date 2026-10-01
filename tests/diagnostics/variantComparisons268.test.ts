// Diagnostics tests: a comparison converts only between typed operands, a
// typed Date or Boolean compared with a string it cannot read, and an object
// read as a condition (issue #268). Every case was measured in Excel 16.0
// (build 20326, 2026-10-01).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

function source(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

function errors(src: string): string[] {
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a Variant compared with a string', () => {
	it('stays quiet: a Variant on either side compares without converting', () => {
		for (const lines of [
			['Dim v As Variant', 'v = 5', 'Main = (v = "")'],
			['Dim v As Variant', 'v = 5', 'Main = (v = "abc")'],
			['Dim v As Variant', 'v = 5', 'Main = (v < "abc")'],
			['Dim v As Variant, s As String', 'v = 5', 's = "abc"', 'Main = (v = s)'],
			['Dim v As Variant, s As String', 'v = 5', 'Main = (v = s)'],
			['Dim v As Variant, w As Variant', 'v = 5', 'w = "abc"', 'Main = (v = w)'],
			['Dim v As Variant', 'v = 5', 'Main = (v = "a" & "b")'],
			['Dim v As Variant', 'v = 5', 'Select Case v', 'Case ""', '    Main = 1', 'Case Else', '    Main = 2', 'End Select'],
			['Dim v As Variant', 'v = 5', 'If v <> "" Then Main = 1 Else Main = 2'],
			['Dim v As Variant', 'v = 5', 'If v <> "" Then', '    Main = 1', 'End If'],
			['Dim v As Variant', 'v = 1.5', 'Main = (v = "abc")'],
			['Dim v As Variant', 'v = "abc"', 'Main = (v = 5)'],
			['Dim v As Variant, n As Long', 'v = "abc"', 'Main = (v = n)'],
			['Dim v As Variant', 'v = "abc"', 'Main = (v < 5)'],
			['Dim v', 'v = 5', 'Main = (v = "abc")'],
			['Dim b As Boolean', 'Main = (b = "True")'],
			['Dim d As Date', 'Main = (d = "1/1/2000")'],
			['Dim n As Long', 'Select Case n', 'Case "1" To "9"', '    Main = 1', 'End Select'],
		]) {
			expect(errors(source(...lines)), lines.join(': ')).toEqual([]);
		}
	});

	it.each([
		[['Dim v As Variant', 'v = 5', 'Main = v + "abc"'], '"abc"', "Operator '+' coerces string literal \"abc\" to a number"],
		[['Dim v As Variant', 'v = 5', 'Main = v - "abc"'], '"abc"', "Operator '-' coerces"],
		[['Dim n As Long', 'Main = (n = "abc")'], '"abc"', "Operator '=' coerces string literal \"abc\" to a number"],
	])('still reports arithmetic, and a typed number compared with a string: %j', (lines, span, message) => {
		const src = source(...lines);
		expectDiagnostic(src, byCode(analyzeModule(src), 'string-arithmetic-coercion'), 'string-arithmetic-coercion', { span, message });
	});
});

describe('a typed Date or Boolean compared with a string it cannot read', () => {
	it.each([
		[['Dim d As Date', 'If d = "" Then Main = 1 Else Main = 2'], '""', "Operator '=' compares string literal \"\" with a Date, which cannot read it"],
		[['Dim d As Date', 'Main = (d = "abc")'], '"abc"', 'with a Date, which cannot read it'],
		[['Dim b As Boolean', 'Main = (b = "yes")'], '"yes"', 'with a Boolean, which cannot read it'],
		[['Dim d As Date', 'Select Case d', 'Case "abc"', '    Main = 1', 'End Select'], '"abc"', 'Case compares string literal "abc" with a Date, which cannot read it'],
		[['Dim n As Long', 'Select Case n', 'Case "a" To "z"', '    Main = 1', 'End Select'], '"a"', 'Case compares string literal "a" with a number'],
		[['Dim n As Long, w As Variant', 'w = "abc"', 'Select Case n', 'Case w', '    Main = 1', 'End Select'], 'w', "Case compares 'w', which holds \"abc\" with a number"],
		[['Dim n As Long, s As String', 's = "abc"', 'Select Case n', 'Case s', '    Main = 1', 'End Select'], 's', "Case compares 's', which holds \"abc\" with a number"],
		[['Dim s As String', 's = "abc"', 'If Not s Then Main = 1'], 's', "Operator 'Not' coerces 's', which holds \"abc\""],
	])('reports %j', (lines, span, message) => {
		const src = source(...lines);
		expectDiagnostic(src, byCode(analyzeModule(src), 'string-arithmetic-coercion'), 'string-arithmetic-coercion', { span, message });
	});
});

describe('an object read as a condition', () => {
	it('reports a Collection, which has no value without an index', () => {
		const src = source('Dim c As New Collection', 'If c Then Main = 1');
		expectDiagnostic(src, byCode(analyzeModule(src), 'object-default-value'), 'object-default-value', { span: 'c', message: "the condition has no value to read. This will raise Run-time error '450'" });
	});

	it('reports one still Nothing', () => {
		const src = source('Dim c As Collection', 'If c Then Main = 1');
		expectDiagnostic(src, byCode(analyzeModule(src), 'object-variable-not-set'), 'object-variable-not-set', { span: 'c', message: "is Nothing when the condition reads its value. This will raise Run-time error '91'" });
	});
});
