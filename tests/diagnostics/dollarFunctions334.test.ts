// Diagnostics tests: a bad Long local reaching a `$` string function or String
// inside a block (issue #334). Every case was measured in Excel 16.0 (build
// 20326, 2026-10-01).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

const CODE = 'runtime-argument-value';

function source(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

describe('a $ function or String with a bad Long local inside a block', () => {
	it('is reported', () => {
		const blocks: Array<[string, string]> = [
			['If True Then', 'End If'],
			// The statements go under a Case: before the first Case they are refused.
			['Select Case 1\n    Case 1', 'End Select'],
			['Do', 'Loop'],
			['For i = 1 To 1', 'Next'],
			['With Application', 'End With'],
		];
		for (const call of ['Left$("abc", n)', 'Right$("abc", n)', 'Mid$("abc", n)', 'Mid$("abc", 1, n)', 'Space$(n)', 'Chr$(n)', 'String(n, "a")', 'String$(n, "a")', 'VBA.Left$("abc", n)']) {
			for (const [head, foot] of blocks) {
				const exit = head === 'Do' ? ['        Exit Do'] : [];
				const src = source('Dim n As Long, i As Long', 'n = -1', head, `        Main = ${call}`, ...exit, foot);
				expectDiagnostic(src, byCode(analyzeModule(src), CODE), CODE, { span: 'n', message: "'5'" });
			}
		}
	});

	it('is reported nested in another call at the top level', () => {
		const src = source('Dim n As Long, m As Variant', 'm = Len(Mid$("abc", n, 1))', 'Main = m');
		expectDiagnostic(src, byCode(analyzeModule(src), CODE), CODE, { span: 'n', message: "'5'" });
	});

	it('stays quiet when the value is in range', () => {
		for (const [value, call] of [['1', 'Left$("abc", n)'], ['2', 'String$(n, "a")'], ['65', 'Chr$(n)']]) {
			const src = source('Dim n As Long', `n = ${value}`, 'If True Then', `    Main = ${call}`, 'End If');
			expect(byCode(analyzeModule(src), CODE), call).toEqual([]);
		}
	});
});
