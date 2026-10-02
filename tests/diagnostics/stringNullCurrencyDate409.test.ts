// Diagnostics tests: String with a Null Character returns Null before Number
// is checked, and a Currency with a Date takes the Date rules (issue #409).
// Measured in Excel 16.0 (build 20326, 2026-10-01).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

function source(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

describe('String with a Null Character', () => {
	it('returns Null whatever Number is', () => {
		for (const lines of [
			['Main = IsNull(String(-1, Null))'],
			['Main = IsNull(String(2147483647, Null))'],
			['Dim v As Variant', 'v = Null', 'Main = IsNull(String(-1, v))'],
		]) {
			expect(byCode(analyzeModule(source(...lines)), 'runtime-argument-value'), lines.join(' : ')).toEqual([]);
		}
		const src = source('Main = String(-1, "a")');
		expectDiagnostic(src, byCode(analyzeModule(src), 'runtime-argument-value'), 'runtime-argument-value', { span: '-1' });
	});
});

describe('a Currency with a Date', () => {
	const setup = ['Dim c As Currency, t As Date', 'c = 922337203685477', 't = #1/2/2000#'];
	it('is a Date for + and -, and a Double for *', () => {
		for (const lines of [
			[...setup, 'Main = VarType(c + t)'],
			['Main = VarType(922337203685477@ + #1/2/2000#)'],
			[...setup, 'Main = VarType(t * c)'],
			[...setup, 'Main = VarType(c * t)'],
		]) {
			expect(byCode(analyzeModule(source(...lines)), 'arithmetic-overflow'), lines.join(' : ')).toEqual([]);
		}
	});

	it('overflows the Date range with the Date first', () => {
		for (const [lines, span] of [
			[[...setup, 'Main = VarType(t + c)'], 't + c'],
			[[...setup, 'Main = VarType(t - c)'], 't - c'],
			[['Dim c As Currency, t As Date', 'c = 3000000', 't = #1/2/2000#', 'Main = VarType(t - c)'], 't - c'],
		] as const) {
			const src = source(...lines);
			expectDiagnostic(src, byCode(analyzeModule(src), 'arithmetic-overflow'), 'arithmetic-overflow', { span, message: 'Date range' });
		}
	});
});
