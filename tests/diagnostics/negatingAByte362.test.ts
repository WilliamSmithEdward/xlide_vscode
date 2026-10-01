// Diagnostics tests: unary minus on a Byte gives an Integer (issue #362).
// Every case was measured in Excel 16.0 (build 20326, 2026-10-01).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

const CODE = 'arithmetic-overflow';

function source(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

describe('negating a Byte', () => {
	it('gives an Integer, which never overflows', () => {
		for (const lines of [
			['Dim b As Byte', 'b = 200', 'Main = -b'],
			['Dim b As Byte', 'b = 200', 'Main = TypeName(-b)'],
			['Dim b As Byte, i As Integer', 'b = 200', 'i = -b', 'Main = i'],
			['Main = -CByte(200)'],
			['Dim b As Byte', 'b = 1', 'Main = -b'],
			['Dim b As Byte', 'b = 200', 'Main = -(-b)'],
		]) {
			expect(byCode(analyzeModule(source(...lines)), CODE), lines.join(': ')).toEqual([]);
		}
	});

	it('still overflows into a Byte, or as an Integer', () => {
		for (const [lines, message] of [
			[['Dim b As Byte, c As Byte', 'b = 200', 'c = -b', 'Main = c'], 'stores -200 in a Byte'],
			[['Dim b As Byte', 'b = 200', 'Main = -b * b'], 'outside the Integer range'],
			[['Dim a As Byte, b As Byte', 'a = 1', 'b = 2', 'Main = a - b'], 'outside the Byte range'],
		] as const) {
			const src = source(...lines);
			expectDiagnostic(src, byCode(analyzeModule(src), CODE), CODE, { message });
		}
	});
});
