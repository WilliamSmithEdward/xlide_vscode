// Diagnostics tests: a guard whose outcome the code makes plain decides
// which statements run (issue #273). Every case was measured in Excel 16.0
// (build 20326, 2026-10-01).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

function source(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => (line.endsWith(':') ? line : `    ${line}`)).join('\n')}\nEnd Function\n`;
}

function errors(src: string): string[] {
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a guard that always leaves', () => {
	it('keeps the code after it from being checked', () => {
		for (const lines of [
			['Dim c As Collection', 'If c Is Nothing Then Exit Function', 'Main = c.Count'],
			['Dim c As Collection', 'If (c Is Nothing) Then Exit Function', 'Main = c.Count'],
			['Dim c As Collection, x As Boolean', 'If c Is Nothing Or x Then Exit Function', 'Main = c.Count'],
			['Dim c As Collection', 'If c Is Nothing Then', '    Main = -1', '    Exit Function', 'End If', 'Main = c.Count'],
			['Dim c As Collection', 'On Error GoTo H', 'If c Is Nothing Then Err.Raise 5', 'Main = c.Count', 'Exit Function', 'H:', 'Main = -2'],
			['Dim c As Collection', 'If c Is Nothing Then GoTo Done', 'Main = c.Count', 'Done:'],
			['Dim c As Collection', 'If c Is Nothing Then Exit Function', 'Set c = Nothing', 'Main = c.Count'],
			['Dim c As Collection, i As Long', 'For i = 1 To 2', '    If c Is Nothing Then Exit For', '    Main = c.Count', 'Next'],
			['Dim d As Long', 'If d = 0 Then Exit Function', 'Main = 10 / d'],
			['Dim d As Long', 'If d = 0 Then', '    Main = 0', '    Exit Function', 'End If', 'Main = 10 / d'],
			['Dim d As Long', 'On Error GoTo H', 'If d = 0 Then Err.Raise 5', 'Main = 10 / d', 'Exit Function', 'H:', 'Main = -2'],
			['Dim a() As Long, n As Long', 'If n = 0 Then Exit Function', 'Main = a(0)'],
		]) {
			expect(errors(source(...lines)), lines.join(': ')).toEqual([]);
		}
	});

	it('still reports when the guard may not leave', () => {
		for (const [lines, code] of [
			[['Dim c As Collection, x As Boolean', 'If c Is Nothing And x Then Exit Function', 'Main = c.Count'], 'object-variable-not-set'],
			[['Dim c As Collection', 'If c Is Nothing Then Main = -1', 'Main = c.Count'], 'object-variable-not-set'],
			[['Dim d As Long', 'If d = 0 Then Main = 1', 'Main = 10 / d'], 'division-by-zero'],
		] as const) {
			const src = source(...lines);
			expectDiagnostic(src, byCode(analyzeModule(src), code), code);
		}
	});
});

describe('an arm a known condition rules out', () => {
	it('is not checked', () => {
		for (const lines of [
			['Dim c As Collection', 'If Not c Is Nothing Then', '    Main = c.Count', 'Else', '    Main = 0', 'End If'],
			['Dim c As Collection', 'If c Is Nothing Then', '    Main = 0', 'Else', '    Main = c.Count', 'End If'],
			['Dim c As Collection', 'If Timer < 0 Then', '    Main = 1', 'ElseIf c Is Nothing Then', '    Main = 0', 'Else', '    Main = c.Count', 'End If'],
			['Dim d As Long', 'Select Case d', 'Case 0', '    Main = 0', 'Case Else', '    Main = 10 / d', 'End Select'],
			['Dim d As Long', 'Select Case d', 'Case Is = 0', '    Main = 0', 'Case Else', '    Main = 10 / d', 'End Select'],
			['Dim d As Long', 'If d = 0 Then', '    Main = 0', 'Else', '    Main = 10 / d', 'End If'],
			['Dim a() As Long, n As Long', 'n = 0', 'If n > 0 Then Main = a(0)', 'Main = 1'],
			['Dim a() As Long, n As Long', 'n = 0', 'If n > 0 Then', '    Main = a(0)', 'End If', 'Main = 1'],
			['Dim a() As Long, n As Long', 'n = 2', 'If n = 1 Then Main = a(0)', 'Main = 1'],
			['Dim s As String', 's = "abc"', 'If IsNumeric(s) Then Main = s + 1', 'Main = 1'],
			['Dim s As String, n As Long', 's = "abc"', 'If n > 0 Then Main = s + 1', 'Main = 1'],
			['Dim s As String, n As Long', 's = "abc"', 'If n > 0 Then', '    Main = s + 1', 'End If', 'Main = 1'],
		]) {
			expect(errors(source(...lines)), lines.join(': ')).toEqual([]);
		}
	});

	it('does not count a Set or a ReDim in it', () => {
		for (const [lines, code] of [
			[['Dim c As Collection', 'If Not c Is Nothing Then Set c = New Collection', 'Main = c.Count'], 'object-variable-not-set'],
			[['Dim c As Collection, n As Long', 'n = 0', 'If n > 0 Then Set c = New Collection', 'Main = c.Count'], 'object-variable-not-set'],
			[['Dim c As Collection, n As Long', 'n = 0', 'If n > 0 Then', '    Set c = New Collection', 'End If', 'Main = c.Count'], 'object-variable-not-set'],
			[['Dim c As Collection, n As Long', 'Select Case n', 'Case 1', '    Set c = New Collection', 'End Select', 'Main = c.Count'], 'object-variable-not-set'],
			[['Dim a() As Long, n As Long', 'If n > 0 Then ReDim a(1)', 'Main = a(0)'], 'unallocated-dynamic-array-access'],
			[['Dim d As Long', 'If d <> 0 Then d = 1', 'Main = 10 / d'], 'division-by-zero'],
			// The same with a label, which the GoTo-following walk takes.
			[['Dim c As Collection', 'If Not c Is Nothing Then Set c = New Collection', 'Here:', 'Main = c.Count'], 'object-variable-not-set'],
		] as const) {
			const src = source(...lines);
			expectDiagnostic(src, byCode(analyzeModule(src), code), code);
		}
	});

	it('still checks the arm that runs', () => {
		for (const [lines, code] of [
			[['Dim d As Long', 'Select Case d', 'Case 1', '    Main = 0', 'Case Else', '    Main = 10 / d', 'End Select'], 'division-by-zero'],
			[['Dim a() As Long, n As Long', 'n = 1', 'If n > 0 Then Main = a(0)', 'Main = 1'], 'unallocated-dynamic-array-access'],
			[['Dim s As String', 's = "abc"', 'If Not IsNumeric(s) Then Main = s + 1', 'Main = 1'], 'string-arithmetic-coercion'],
		] as const) {
			const src = source(...lines);
			expectDiagnostic(src, byCode(analyzeModule(src), code), code);
		}
	});
});

describe('an If arm that always leaves', () => {
	it('takes no part in the state after the block', () => {
		for (const lines of [
			['Dim c As Collection', 'If Timer < 0 Then', '    Exit Function', 'Else', '    Set c = New Collection', 'End If', 'Main = c.Count'],
			['Dim c As Collection', 'If Timer > 0 Then', '    Set c = New Collection', 'Else', '    Exit Function', 'End If', 'Main = c.Count'],
			['Dim c As Collection', 'If c Is Nothing Then Set c = New Collection', 'Main = c.Count'],
			['Dim c As Collection', 'Dim n As Long', 'Select Case n', 'Case 0', '    Set c = New Collection', 'End Select', 'Main = c.Count'],
		]) {
			expect(errors(source(...lines)), lines.join(': ')).toEqual([]);
		}
	});

	it('leaves only the path that skips the block', () => {
		for (const lines of [
			['Dim c As Collection', 'If Timer < 0 Then', '    Set c = New Collection', '    Exit Function', 'End If', 'Main = c.Count'],
			['Dim c As Collection', 'If Timer < 0 Then', '    If c Is Nothing Then Exit Function', 'End If', 'Main = c.Count'],
		]) {
			const src = source(...lines);
			expectDiagnostic(src, byCode(analyzeModule(src), 'object-variable-not-set'), 'object-variable-not-set', { span: 'c' });
		}
	});
});

describe('Err.Raise under On Error Resume Next', () => {
	it('does not leave', () => {
		const src = source('Dim d As Long', 'On Error Resume Next', 'If d = 0 Then Err.Raise 5', 'Main = 10 / d');
		expectDiagnostic(src, byCode(analyzeModule(src), 'division-by-zero'), 'division-by-zero');
	});
});
