// Diagnostics tests: statement forms and literals the VBE refuses (issue
// #369). Every case was measured in Excel 16.0 (build 20326, 2026-10-01),
// compiled with the VBE's Debug > Compile.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

function source(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

function errors(src: string): string[] {
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('statement forms the VBE refuses', () => {
	it.each([
		[['Dim a As Boolean, b As Boolean', 'If a Then', '    Main = 1', 'Else If b Then', '    Main = 2', 'End If'], 'Else If b Then', 'Syntax error'],
		[['Dim x As Long', 'x += 1', 'Main = x'], '+=', 'Syntax error'],
		[['Dim x As Long', 'x -= 1', 'Main = x'], '-=', 'Syntax error'],
		[['Dim c As New Collection', 'c.Count += 1', 'Main = 1'], '+=', 'c.Count = c.Count +'],
		[['Dim n As Long', 'While n < 3', '    n = n + 1', '    If n = 2 Then Exit While', 'Wend', 'Main = n'], 'Exit While', 'Syntax error'],
		[['Call Debug.Print("a")', 'Main = 1'], 'Call Debug.Print', 'Syntax error'],
		[['Call Debug.Assert(True)', 'Main = 1'], 'Call Debug.Assert', 'Syntax error'],
		[['Dim n As Long', 'Select Case n', 'Case Else', '    Main = 0', 'Case 1', '    Main = 1', 'End Select'], 'Case 1', 'Case without Select Case'],
		[['Dim n As Long', 'Do Until n > 3', '    n = n + 1', 'Loop While n < 10', 'Main = n'], 'Loop While', 'Loop without Do'],
		[['Dim n As Long', 'Do', '    n = n + 1', 'Loop While n < 3 Until n > 5', 'Main = n'], 'While n < 3 Until', 'Syntax error'],
		[['Dim n As Long', 'Select Case n', 'Case Is = 1 To 5', '    Main = 1', 'End Select'], 'Is = 1 To', 'Syntax error'],
		[['Dim s As String', 'Select Case s', 'Case Is Like "a*"', '    Main = 1', 'End Select'], 'Is Like', 'Syntax error'],
	])('reports %j', (lines, span, error) => {
		const src = source(...lines);
		expectDiagnostic(src, byCode(analyzeModule(src), 'malformed-statement'), 'malformed-statement', { span, message: error });
	});

	it('leaves *= and ^= to the operator-run check, once each', () => {
		for (const [lines, span] of [
			[['Dim x As Long', 'x *= 2', 'Main = x'], '*='],
			[['Dim x As Double', 'x = 2', 'x ^= 2', 'Main = x'], '^='],
		] as const) {
			const src = source(...lines);
			expectDiagnostic(src, byCode(analyzeModule(src), 'invalid-expression-syntax'), 'invalid-expression-syntax', { span });
			expect(byCode(analyzeModule(src), 'malformed-statement')).toEqual([]);
		}
	});

	it('leaves the forms that compile alone', () => {
		for (const lines of [
			['Dim a As Boolean, b As Boolean', 'If a Then', '    Main = 1', 'Else If b Then Main = 2', 'End If'],
			['Dim a As Boolean, b As Boolean', 'If a Then', '    Main = 1', 'ElseIf b Then', '    Main = 2', 'End If'],
			['Dim x&', 'x&=5', 'Main = x'],
			['Dim x^', 'x^=5', 'Main = x'],
			['Dim x&', 'x& = 5', 'Main = x'],
			['Dim n As Long', 'Select Case n', 'Case Is > 1, 5 To 7', '    Main = 1', 'End Select'],
			['Dim n As Long', 'Select Case n', 'Case Is >= 1', '    Main = 1', 'End Select'],
			['Dim n As Long, m As Long', 'Select Case n', 'Case 0', '    Select Case m', '    Case Else', '        Main = 1', '    End Select', 'Case 1', '    Main = 2', 'End Select'],
			['Dim n As Long', 'Do', '    n = n + 1', 'Loop Until n > 3', 'Main = n'],
			['Dim n As Long', 'Do While n < 3', '    n = n + 1', 'Loop', 'Main = n'],
			['Dim n As Long, m As Long', 'Do While n < 2', '    Do', '        m = m + 1', '    Loop Until m > 2', '    n = n + 1', 'Loop', 'Main = m'],
			['Dim n As Long', 'Do While n < 3', '    n = n + 1', '    If n = 2 Then Exit Do', 'Loop', 'Main = n'],
			['Debug.Print "a"', 'Main = 1'],
		]) {
			expect(errors(source(...lines)), lines.join(': ')).toEqual([]);
		}
	});
});

describe('hex and octal literals the VBE refuses', () => {
	it.each([
		['&H100000000', 'wider than 32 bits'],
		['&O40000000000', 'wider than 32 bits'],
		['&HFF#', 'Double suffix'],
		['&HFF!', 'Single suffix'],
	])('reports %s', (literal, message) => {
		const src = source('Dim v As Variant', `v = ${literal}`, 'Main = v');
		expectDiagnostic(src, byCode(analyzeModule(src), 'suffixed-literal-overflow'), 'suffixed-literal-overflow', { span: literal, message });
	});

	it('leaves 32 bits and the & and ^ suffixes alone', () => {
		for (const literal of ['&HFFFFFFFF', '&O37777777777', '&HFF&', '&H100000000^']) {
			expect(errors(source('Dim v As Variant', `v = ${literal}`, 'Main = v')), literal).toEqual([]);
		}
	});
});
