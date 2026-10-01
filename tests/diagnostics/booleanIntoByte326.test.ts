// Diagnostics tests: a Boolean stored in a Byte is 255 or 0 (issue #326).
// Every case was measured in Excel 16.0 (build 20326, 2026-10-01).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

function source(lines: readonly string[], extra = ''): string {
	return `Option Explicit\n${extra}Function Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

function errors(src: string): string[] {
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a Boolean stored in a Byte', () => {
	it('runs: True is 255 and False 0', () => {
		for (const lines of [
			['Dim b As Byte', 'b = True', 'Main = b'],
			['Dim b As Byte', 'b = Not False', 'Main = b'],
			['Dim b As Byte', 'b = Not Not True', 'Main = b'],
			['Dim b As Byte', 'b = (True)', 'Main = b'],
			['Dim b As Byte', 'b = False', 'Main = b'],
			['Dim b As Byte', 'b = Not True', 'Main = b'],
			['Const T As Boolean = True', 'Dim b As Byte', 'b = T', 'Main = b'],
			['Const T As Boolean = 1', 'Dim b As Byte', 'b = T', 'Main = b'],
			['Const B As Byte = True', 'Main = B'],
			['Dim f(2) As Byte', 'f(1) = True', 'Main = f(1)'],
			['Main = CByte(True)'],
			['Main = CByte(Not False)'],
			['Dim b As Byte', 'b = -True', 'Main = b'],
			['Dim i As Integer', 'i = True', 'Main = i'],
			['Main = CInt(True)'],
			['Dim b As Byte', 'b = True', 'b = b - 1', 'Main = b'],
		]) {
			expect(errors(source(lines)), lines.join(': ')).toEqual([]);
		}
		expect(errors(source(['Main = RetB()'], 'Private Function RetB() As Byte\n    RetB = True\nEnd Function\n'))).toEqual([]);
	});

	it('still overflows once arithmetic makes it an Integer -1', () => {
		for (const [lines, message] of [
			[['Dim b As Byte', 'b = True + 0', 'Main = b'], 'stores -1'],
			[['Dim b As Byte', 'b = Not 0', 'Main = b'], 'stores -1'],
			[['Main = CByte(True + 0)'], 'CByte(-1)'],
			// A Boolean Const of 1 is True; a Byte that took False holds 0.
			[['Const T As Boolean = 1', 'Dim b As Byte', 'b = T + 0', 'Main = b'], 'stores -1'],
			[['Dim b As Byte', 'b = False', 'b = b - 1', 'Main = b'], 'stores -1'],
		] as const) {
			const src = source(lines);
			expectDiagnostic(src, byCode(analyzeModule(src), 'arithmetic-overflow'), 'arithmetic-overflow', { message: [message, "'6'"] });
		}
	});

	it('reads vbTrue, a VbTriState -1, as an overflow', () => {
		const src = source(['Dim b As Byte', 'b = vbTrue', 'Main = b']);
		expectDiagnostic(src, byCode(analyzeModule(src), 'assignment-type-mismatch'), 'assignment-type-mismatch', { span: 'vbTrue', message: ["stores vbTrue, which is -1, in a Byte", "'6': Overflow"] });
	});
});
