// Diagnostics tests: True, False and Empty passed to a built-in are -1, 0
// and 0, and 1E10 overflows a Long parameter (issue #434). Measured in
// Excel 16.0 64-bit (build 20430, 2026-10-02) through pyVBAharness.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(call: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    Main = ${call}\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => `${diag.code}: ${diag.message}`);
}

describe('True, False and Empty as a built-in argument', () => {
	it('raise 5 where -1 or 0 does', () => {
		for (const call of [
			'Left("abc", True)', 'Space(True)', 'String(True, "a")', 'Mid("abc", True)', 'Mid("abc", False)', 'Mid("abc", Empty)',
			'Mid("abc", 1, True)', 'MonthName(True)', 'MonthName(False)', 'WeekdayName(Empty)', 'StrComp("a", "b", True)',
			'StrConv("a", True)', 'InStr(False, "abc", "b")', 'InStrRev("abc", "b", False)', 'Replace("abc", "b", "x", Empty)',
			'Weekday(Now, True)', 'Chr(True)',
		]) {
			expect(errors(call), call).toEqual([expect.stringContaining("'5'")]);
		}
	});

	it('run where -1 does', () => {
		expect(errors('InStrRev("abc", "b", True)')).toEqual([]);
	});
});

describe('1E10 as a whole-number built-in argument', () => {
	it('overflows', () => {
		for (const call of ['MonthName(1E10)', 'WeekdayName(1E10)', 'StrComp("a", "b", 1E10)', 'StrConv("a", 1E10)', 'InStrRev("abc", "b", 1E10)', 'Weekday(Now, 1E10)', 'Chr(1E10)']) {
			expect(errors(call), call).toEqual([expect.stringContaining("'6'")]);
		}
		expect(errors('Left("abc", 1E10)')).toHaveLength(1);
	});
});
