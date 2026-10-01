// Diagnostics tests: `number + Date` past the Date range does not raise
// (issue #330), but most uses of the result do (issue #405). Every case was
// measured in Excel 16.0 (build 20326, 2026-10-01).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

function source(...lines: string[]): string {
	const setup = ['Dim i As Integer, d As Date', 'i = 32767', 'd = #12/31/9999#'];
	return `Option Explicit\nFunction Main() As Variant\n${[...setup, ...lines].map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

function errors(src: string): string[] {
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => `${diag.code}: ${diag.message}`);
}

describe('a Date past the range made by number + Date', () => {
	it('raises 5 where it is read as text or as a date', () => {
		for (const [line, span] of [
			['Main = CStr(i + d)', 'i + d'],
			['Main = Year(i + d)', 'i + d'],
			['Main = "x" & (i + d)', '(i + d)'],
			['Main = (i + d) & "x"', '(i + d)'],
			['Main = Format(i + d, "yyyy")', 'i + d'],
			['Main = DateAdd("d", 1, i + d)', 'i + d'],
			['Main = Str(i + d)', 'i + d'],
			['Main = CStr(0 - d)', '0 - d'],
			['Main = CStr(1 + #12/31/9999#)', '1 + #12/31/9999#'],
			['Debug.Print i + d', 'i + d'],
			['Main = CStr(CDate(i + d))', 'CDate(i + d)'],
		] as const) {
			const src = source(line);
			expectDiagnostic(src, byCode(analyzeModule(src), 'runtime-argument-value'), 'runtime-argument-value', { span, message: ["'5'", 'outside the Date range'] });
		}
		const stored = source('Dim t As Date', 't = i + d', 'Main = CStr(t)');
		expectDiagnostic(stored, byCode(analyzeModule(stored), 'runtime-argument-value'), 'runtime-argument-value', { span: 't', message: "'5'" });
	});

	it('raises 6 where the next operation needs it in range', () => {
		for (const [line, span] of [
			['Dim n As Integer: n = i + d', 'i + d'],
			['Main = Int(i + d)', 'Int(i + d)'],
			['Main = CInt(i + d)', 'CInt(i + d)'],
			['Main = (i + d) + 1', '(i + d) + 1'],
			['Main = -(i + d)', '-(i + d)'],
		] as const) {
			const src = source(...line.split(': '));
			expectDiagnostic(src, byCode(analyzeModule(src), 'arithmetic-overflow'), 'arithmetic-overflow', { span, message: "'6'" });
		}
	});

	it('runs where it is stored, converted to a number, compared or typed', () => {
		for (const lines of [
			['Main = i + d'],
			['Dim t As Date', 't = i + d', 'Main = CDbl(t)'],
			['Dim n As Long', 'n = i + d', 'Main = n'],
			['Dim s As Single', 's = i + d', 'Main = s'],
			['Main = CDbl(i + d)'],
			['Main = CLng(i + d)'],
			['Main = CCur(i + d)'],
			['Main = CDate(i + d)'],
			['Main = IsDate(i + d)'],
			['Main = TypeName(i + d)'],
			['Main = (i + d) > d'],
			['Main = 1 + (i + d)'],
			['Main = (i + d) - d'],
		]) {
			expect(errors(source(...lines)), lines.join(' : ')).toEqual([]);
		}
	});
});
