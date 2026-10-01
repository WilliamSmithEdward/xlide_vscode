// Diagnostics tests: the object and dynamic-array state rules follow GoTo
// (issue #271). Every case was measured in Excel 16.0 (build 20326,
// 2026-10-01).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

function source(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => (line.endsWith(':') ? line : `    ${line}`)).join('\n')}\nEnd Function\n`;
}

function errors(src: string): string[] {
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a backward GoTo into code after the Set', () => {
	it('stays quiet when every way into the label has set it', () => {
		for (const lines of [
			['Dim c As Collection', 'GoTo Setup', 'Use:', 'c.Add 1', 'Main = c.Count', 'Exit Function', 'Setup:', 'Set c = New Collection', 'GoTo Use'],
			['Dim c As Collection', 'GoTo Setup', 'Use:', 'With c', '    Main = .Count', 'End With', 'Exit Function', 'Setup:', 'Set c = New Collection', 'GoTo Use'],
			['Dim c As Collection, x As Variant', 'GoTo Setup', 'Use:', 'For Each x In c', 'Next', 'Main = 1', 'Exit Function', 'Setup:', 'Set c = New Collection', 'GoTo Use'],
			['Dim c As Collection', 'Set c = Nothing', 'GoTo Setup', 'Use:', 'Main = c.Count', 'Exit Function', 'Setup:', 'Set c = New Collection', 'GoTo Use'],
			['Dim c As Collection', 'On Error GoTo Setup', 'Err.Raise 5', 'Use:', 'Main = c.Count', 'Exit Function', 'Setup:', 'Set c = New Collection', 'Resume Use'],
			['Dim a() As Long', 'GoTo Setup', 'Use:', 'Main = UBound(a)', 'Exit Function', 'Setup:', 'ReDim a(2)', 'GoTo Use'],
			['Dim a() As Long', 'ReDim a(2)', 'Erase a', 'GoTo Setup', 'Use:', 'Main = UBound(a)', 'Exit Function', 'Setup:', 'ReDim a(2)', 'GoTo Use'],
			['Dim c As Collection, n As Long', 'If n = 1 Then GoTo Use', 'Set c = New Collection', 'Use:', 'Main = c.Count'],
			['Dim c As Collection, tries As Long', 'Retry:', 'tries = tries + 1', 'If tries < 2 Then', '    Set c = New Collection', '    GoTo Retry', 'End If', 'Main = c.Count'],
			['Dim c As Collection', 'Set c = New Collection', 'Here:', 'Main = c.Count'],
			// Code nothing reaches is not checked.
			['Dim c As Collection', 'GoTo Done', 'Main = c.Count', 'Done:', 'Main = 1'],
			// An If arm with no Else may not run: what follows merges both ways.
			['Dim c As Collection, n As Long', 'Set c = New Collection', 'GoTo Go', 'Go:', 'If n = 1 Then', '    Set c = Nothing', 'End If', 'Main = c.Count'],
			// A GoTo inside a block enters its label with nothing known.
			['Dim c As Collection, n As Long', 'n = 1', 'If n = 1 Then GoTo Skip', 'Use:', 'Main = c.Count', 'Exit Function', 'Skip:', 'If n = 1 Then', '    Set c = New Collection', '    GoTo Use', 'End If'],
		]) {
			expect(errors(source(...lines)), lines.join(': ')).toEqual([]);
		}
	});
});

describe('a forward GoTo past the Set', () => {
	it.each([
		[['Dim c As Collection', 'GoTo Use', 'Setup:', 'Set c = New Collection', 'Use:', 'Main = c.Count'], 'object-variable-not-set', 'c'],
		[['Dim c As Collection', 'GoTo Use', 'Set c = New Collection', 'Use:', 'Main = c.Count'], 'object-variable-not-set', 'c'],
		[['Dim a() As Long', 'GoTo Use', 'Setup:', 'ReDim a(2)', 'Use:', 'Main = UBound(a)'], 'unallocated-dynamic-array-access', 'a'],
		[['Dim c As Collection', 'GoTo Use', 'Use:', 'Main = c.Count'], 'object-variable-not-set', 'c'],
		// A GoTo in a loop leaves with what the loop did not touch.
		[['Dim c As Collection, i As Long', 'For i = 1 To 2', '    If i = 2 Then GoTo Use', 'Next', 'Use:', 'Main = c.Count'], 'object-variable-not-set', 'c'],
		[['Dim c As Collection, i As Long', 'For i = 1 To 2', '    If i = 2 Then GoTo Use', 'Next', 'Exit Function', 'Use:', 'Main = c.Count'], 'object-variable-not-set', 'c'],
		// A GoTo in a single-line If may not run, so the next line is reached.
		[['Dim c As Collection, n As Long', 'If n = 1 Then GoTo Done', 'Main = c.Count', 'Done:'], 'object-variable-not-set', 'c'],
	])('reports %j', (lines, code, span) => {
		const src = source(...lines);
		expectDiagnostic(src, byCode(analyzeModule(src), code), code, { span });
	});

	it('still reports the plain case once', () => {
		expect(byCode(analyzeModule(source('Dim c As Collection', 'Main = c.Count')), 'object-variable-not-set')).toHaveLength(1);
	});
});
