// Diagnostics tests: a literal that raises, in code a known guard keeps from
// running, is not reported (issue #430). Measured in Excel 16.0 (build 20326,
// 2026-10-02).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const LINES = ['y = "abc" + 1', 'y = Sqr(-1)', 'y = 32767 + 1', 'y = Asc("")', 'y = CLng("abc")', 'y = Len(Mid("abc", 0))'];
const SETUP = ['Dim x As Long, y As Double, i As Long', 'x = 2'];

function errors(...lines: string[]): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n${[...SETUP, ...lines].map((line) => `    ${line}`).join('\n')}\n    Main = 1\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a literal that raises, under a guard decided against it', () => {
	it('is not reported in the Then, the block, the block Else or the one-line Else', () => {
		for (const line of LINES) {
			for (const lines of [
				[`If x = 1 Then ${line}`],
				['If x = 1 Then', `    ${line}`, 'End If'],
				['If x = 2 Then', '    y = 0', 'Else', `    ${line}`, 'End If'],
				[`If x = 2 Then y = 0 Else ${line}`],
			]) {
				expect(errors(...lines), lines.join(' / ')).toEqual([]);
			}
		}
	});

	it('is not reported after Exit Function or in a loop of no pass', () => {
		expect(errors('Main = 1', 'Exit Function', 'y = Sqr(-1)')).toEqual([]);
		expect(errors('For i = 1 To 0', '    y = 32767 + 1', 'Next')).toEqual([]);
	});

	it('is reported where the guard lets it run', () => {
		expect(errors('If x = 2 Then y = Sqr(-1)')).toEqual(['runtime-argument-value']);
		expect(errors('If x = 2 Then', '    y = 32767 + 1', 'End If')).toEqual(['arithmetic-overflow']);
		expect(errors('If x = 1 Then y = 0 Else y = Sqr(-1)')).toEqual(['runtime-argument-value']);
	});
});
