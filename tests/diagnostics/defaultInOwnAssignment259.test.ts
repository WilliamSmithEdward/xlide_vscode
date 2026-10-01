// Diagnostics tests: a local's declared default, read by the statement that
// first assigns it (issue #259). Every case was measured in Excel 16.0
// (build 20326, 2026-10-01).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { straightLineAssignments } from '../../src/analyzer/diagnostics/straightLineValues';
import { rawExpressionTokens } from '../../src/analyzer/diagnostics/walker';
import type { ProcedureNode } from '../../src/analyzer/parser/nodes';
import { parseModule } from '../../src/analyzer/parser/parseModule';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

function source(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\nPrivate Sub Fill(ByRef v As Double)\n    v = 4\nEnd Sub\nPrivate Function Fill2(ByVal a As Long, ByRef v As Double) As Long\n    v = 4\nEnd Function\n`;
}

function errors(src: string): string[] {
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a local read by its own first assignment', () => {
	it.each([
		[['Dim x As Long', 'x = 1 / x'], '/'],
		[['Dim x As Long', 'x = 10 \\ x'], '\\'],
		[['Dim x As Long', 'x = 10 Mod x'], 'Mod'],
		[['Dim x As Long', 'x = x + 1 / x'], '/'],
		[['Dim x As Double', 'x = 1 / x'], '/'],
	])('holds its default 0: %j', (lines, operator) => {
		const src = source(...lines, 'Main = 1');
		expectDiagnostic(src, byCode(analyzeModule(src), 'division-by-zero'), 'division-by-zero', { span: 'x', message: `'${operator}' with a zero divisor` });
	});

	it('gives Mid$ a start of 0', () => {
		const src = source('Dim n As Long', 'n = Len(Mid$("abc", n, 1))', 'Main = n');
		expectDiagnostic(src, byCode(analyzeModule(src), 'runtime-argument-value'), 'runtime-argument-value', { span: 'n', message: "'Start' of 'Mid$' is 0" });
	});

	it('holds it before a later assignment too', () => {
		const src = source('Dim d As Long, x As Double', 'x = 10 / d', 'd = 2', 'Main = x');
		expect(errors(src)).toContain('division-by-zero');
	});

	it('holds a String\'s "" the same way', () => {
		const src = source('Dim s As String, n As Long', 'n = CLng(s)', 's = "5"', 'Main = n');
		expectDiagnostic(src, byCode(analyzeModule(src), 'runtime-conversion-value'), 'runtime-conversion-value', { span: 's', message: 'holds "" here' });
	});

	it('walks a body afresh for each start', () => {
		const src = 'Sub T()\n    x = 1\nEnd Sub\n';
		const body = (parseModule(src).members[0] as ProcedureNode).body;
		const start = new Map([['y', rawExpressionTokens('0')]]);
		const plain = straightLineAssignments(src, body, undefined);
		const seeded = straightLineAssignments(src, body, undefined, start);
		expect(plain.get(body[0])?.has('y')).not.toBe(true);
		expect(seeded.get(body[0])?.has('y')).toBe(true);
	});

	it('stays quiet once assigned, in a loop, after a label, or after a ByRef pass', () => {
		for (const lines of [
			['Dim x As Long', 'x = 2', 'x = 1 / x', 'Main = x'],
			['Dim x As Double, i As Long', 'For i = 1 To 2', '    If i = 2 Then x = 1 / x', '    x = 4', 'Next i', 'Main = x'],
			['Dim x As Double', 'GoTo Skip', 'Back:', 'x = 1 / x', 'Main = x', 'Exit Function', 'Skip:', 'x = 2', 'GoTo Back'],
			['Dim x As Double', 'Fill x', 'x = 1 / x', 'Main = x'],
			['Static x As Double', 'x = x + 1', 'Main = 1 / x'],
			['Dim s As String', 's = s & "a"', 'Main = s'],
			['Dim n As Double', 'If Fill2(1, _', '        n) = 0 Then', '    Main = -1', 'End If', 'Main = 10 / n'],
			['Dim n As Double', 'Do While Fill2(1, _', '              n) = 0', '    Exit Do', 'Loop', 'Main = 10 / n'],
			['Dim n As Double', 'If Fill2(1, n) = 0 Then', '    Main = -1', 'End If', 'Main = 10 / n'],
			['Dim n As Double', 'If False Then', 'ElseIf Fill2(1, n) = 0 Then', 'End If', 'Main = 10 / n'],
			['Dim n As Double', 'Do', '    Main = 1', 'Loop While Fill2(1, n) = 1', 'Main = 10 / n'],
		]) {
			expect(errors(source(...lines)), lines.join(': ')).toEqual([]);
		}
	});
});
