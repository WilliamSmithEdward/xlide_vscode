// Diagnostics tests: a number plus or minus a Date never overflows; a Date
// plus or minus a number does (issue #330). Every case was measured in
// Excel 16.0 (build 20326, 2026-10-01).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

const CODE = 'arithmetic-overflow';
const SETUP = ['Dim i As Integer, z As Long, d As Date', 'i = 32767', 'z = 0', 'd = #12/31/9999#'];

function source(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

describe('+ and - between a number and a Date', () => {
	it('never overflow with the number first', () => {
		for (const lines of [
			[...SETUP, 'Main = CDbl(i + d)'],
			[...SETUP, 'Main = TypeName(i + d)'],
			[...SETUP, 'Dim r As Date', 'r = i + d', 'Main = CDbl(r)'],
			['Main = CDbl(1 + #12/31/9999#)'],
			[...SETUP, 'Main = CDbl(z - d)'],
			['Dim d As Date', 'd = #1/1/100#', 'Main = CDbl(1 - d)'],
		]) {
			expect(byCode(analyzeModule(source(...lines)), CODE), lines.join(': ')).toEqual([]);
		}
	});

	it('overflow with the Date first, or two Dates', () => {
		for (const lines of [
			[...SETUP, 'Main = CDbl(d + i)'],
			[...SETUP, 'Main = CDbl(d + 1)'],
			[...SETUP, 'Main = CDbl(d + d)'],
			['Dim d As Date', 'd = #1/1/100#', 'Main = CDbl(d - 1)'],
		]) {
			const src = source(...lines);
			expectDiagnostic(src, byCode(analyzeModule(src), CODE), CODE, { message: 'outside the Date range' });
		}
	});
});
