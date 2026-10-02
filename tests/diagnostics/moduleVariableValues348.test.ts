// Diagnostics tests: a module variable written just before its use (issue
// #348). Each raising sample was measured through pyVBAharness on
// 2026-10-02 in Excel 16.0 (build 20326); each quiet one runs there.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode } from '../helpers/diagnostics';

function module(decl: string, ...lines: string[]): string {
	return `Option Explicit\n${decl}\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

describe('a module variable written just before its use (issue #348)', () => {
	it('follows the value into the rule a local would reach', () => {
		const cases: Array<[string, string[], string]> = [
			['Private m As Integer', ['m = 0', 'Main = CStr(Mid("abc", m))'], 'runtime-argument-value'],
			['Private k As String', ['k = "xyz"', 'Main = CStr(DateValue(k))'], 'runtime-conversion-value'],
			['Private m As Integer', ['m = 1', 'Dim c As New Collection', 'Main = c(m)'], 'collection-index-out-of-range'],
			['Private d As Long', ['d = 0', 'Main = 10 / d'], 'division-by-zero'],
			['Private d As Variant', ['d = 0', 'Main = 10 / d'], 'division-by-zero'],
			['Public d As Double', ['d = 0', 'Debug.Print "x"', 'Main = 10 / d'], 'division-by-zero'],
			['Dim a(3) As Long\nPrivate j As Long', ['j = 5', 'Main = a(j)'], 'array-subscript-out-of-bounds'],
		];
		for (const [decl, lines, code] of cases) {
			expect(byCode(analyzeModule(module(decl, ...lines)), code), lines.join(' / ')).toHaveLength(1);
		}
	});

	it('stays quiet when code between may write it, or the value changed', () => {
		const sources = [
			module('Private d As Long', 'd = 0', 'SetD', 'Main = 10 / d') + 'Sub SetD()\n    d = 2\nEnd Sub\n',
			module('Private d As Long', 'd = 0', 'Main = 10 / Twice() + 10 / d') + 'Function Twice() As Long\n    d = 5\n    Twice = 1\nEnd Function\n',
			module('Private d As Long', 'd = 0', 'd = 2', 'Main = 10 / d'),
			module('Private d As Long', 'd = 0', 'If Rnd() > 2 Then d = 1', 'Main = 10 / d'),
			'Option Explicit\nPrivate d As Long\nFunction Main() As Variant\n    Main = Helper(d)\nEnd Function\n'
				+ 'Function Helper(ByRef x As Long) As Variant\n    d = 0\n    x = 2\n    Helper = 10 / d\nEnd Function\n',
		];
		for (const src of sources) {
			expect(byCode(analyzeModule(src), 'division-by-zero'), src).toHaveLength(0);
		}
	});
});
