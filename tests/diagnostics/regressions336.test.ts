// Diagnostics tests: the regressions issue #336 lists. Every case was measured
// in Excel 16.0 (build 20326, 2026-10-01).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

function source(lines: readonly string[], extra = '', head = ''): string {
	return `Option Explicit\n${head}${extra}Function Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

function errors(src: string): string[] {
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a variable named Property', () => {
	it('is a variable', () => {
		expect(errors(source(['Dim Property As Long', 'Property = 1', 'Main = Property + 1']))).toEqual([]);
		const undeclared = 'Function Main() As Variant\n    Main = CStr(Property)\nEnd Function\n';
		expect(byCode(analyzeModule(undeclared), 'reserved-keyword-in-expression')).toEqual([]);
	});
});

describe('a number in a String into a Date', () => {
	it('is that day', () => {
		for (const value of ['"&HFF"', '"&O17"', '"&HFFFF"', '"2958465"', '"2958465.5"', '"-657434"']) {
			expect(errors(source(['Dim x As Date', `x = ${value}`, 'Main = x'])), value).toEqual([]);
		}
		expect(errors(source(['Main = Echo("&HFF")'], 'Private Function Echo(ByVal x As Date) As Date\n    Echo = x\nEnd Function\n'))).toEqual([]);
	});

	it('raises past the Date range', () => {
		for (const value of ['"&H7FFFFFFF"', '"3000000"', '"2958466"', '"-657435"', '"abc"']) {
			const src = source(['Dim x As Date', `x = ${value}`, 'Main = x']);
			expectDiagnostic(src, byCode(analyzeModule(src), 'assignment-type-mismatch'), 'assignment-type-mismatch', { message: "'13'" });
		}
	});
});

describe('a For Each control variable after the loop', () => {
	it('is Nothing when the loop ends, over an empty collection too', () => {
		for (const lines of [
			['Dim ws As Worksheet', 'For Each ws In Worksheets', 'Next ws', 'Main = ws.Name'],
			['Dim ws As Object, c As New Collection', 'Set ws = Worksheets(1)', 'For Each ws In c', 'Next', 'Main = ws.Name'],
			['Dim ws As Object, c As New Collection', 'For Each ws In c', 'Next', 'Main = ws.Name'],
		]) {
			const src = source(lines);
			expectDiagnostic(src, byCode(analyzeModule(src), 'object-variable-not-set'), 'object-variable-not-set', { span: 'ws' });
		}
	});

	it('is not known when the body can leave the loop', () => {
		for (const lines of [
			['Dim ws As Worksheet', 'For Each ws In Worksheets', '    Exit For', 'Next ws', 'Main = ws.Name'],
			['Dim ws As Worksheet', 'For Each ws In Worksheets', '    GoTo Done', 'Next ws', 'Done:', 'Main = ws.Name'],
		]) {
			expect(byCode(analyzeModule(source(lines)), 'object-variable-not-set'), lines.join(': ')).toEqual([]);
		}
	});
});

describe('Chr of a Double or Currency local', () => {
	it('overflows past the Long range', () => {
		for (const [type, value] of [['Double', '1E300'], ['Currency', '922337203685477@']]) {
			const src = source([`Dim a As ${type}`, `a = ${value}`, 'Main = Chr(a)']);
			expectDiagnostic(src, byCode(analyzeModule(src), 'runtime-argument-value'), 'runtime-argument-value', { span: 'a', message: ["'6': Overflow", value === '1E300' ? 'is 1E+300' : 'is 922337203685477'] });
		}
		const src = source(['Dim a As Double', 'a = 300', 'Main = Chr(a)']);
		expectDiagnostic(src, byCode(analyzeModule(src), 'runtime-argument-value'), 'runtime-argument-value', { message: "'5'" });
		expect(errors(source(['Dim a As Double', 'a = 65.4', 'Main = Chr(a)']))).toEqual([]);
	});
});

describe('Like with a * right before an unclosed [', () => {
	it('reaches the list with any character left', () => {
		for (const expr of ['"abc" Like "a*["', '"b" Like "*["', '"abcd" Like "*["', '"b" Like "*[a"', '"ab" Like "*[z-a]"', '"b" Like "*[a]["', '"bx" Like "*x["']) {
			const src = source([`Main = CStr(${expr})`]);
			expectDiagnostic(src, byCode(analyzeModule(src), 'runtime-argument-value'), 'runtime-argument-value', { message: "'93'" });
		}
	});

	it('stays quiet with nothing left', () => {
		for (const expr of ['"a" Like "a*["', '"" Like "*["', '"x" Like "?["', '"b" Like "*[a]"']) {
			expect(byCode(analyzeModule(source([`Main = CStr(${expr})`])), 'runtime-argument-value'), expr).toEqual([]);
		}
	});
});
