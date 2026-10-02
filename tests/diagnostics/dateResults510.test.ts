// Diagnostics tests: date results handed to the next call (issue #510). Each
// sample was measured through pyVBAharness on 2026-10-02 in Excel 16.0
// (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(expr: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    Main = ${expr}\nEnd Function\n`;
	return analyzeModule(src).filter((d) => d.severity === 'error').map((d) => /Run-time error '(\d+)'/.exec(d.message)?.[1] ?? d.code ?? '');
}

describe('date results handed on (issue #510)', () => {
	it('reads DateValue, CDate, Year, Month, Day and DateDiff of known dates', () => {
		const cases: Array<[string, string]> = [
			['DateAdd("d", 1, DateValue("12/31/9999"))', '5'],
			['DateSerial(9999, DateDiff("d", #1/1/1900#, #3/15/2023#), 1)', '6'],
			['TimeSerial(0, DateDiff("s", #1/1/100#, #12/31/9999#) - 366, 0)', '6'],
			['DateAdd("w", Year(#6/15/5000#), DateValue("12/31/9999"))', '5'],
			['DateAdd("yyyy", 1, CDate("9999-06-01"))', '5'],
			['DateAdd("m", Month(#12/1/2000#), DateSerial(9999, 1, 1))', '5'],
			['TimeSerial(DateDiff("h", #1/1/2000#, #1/1/2010#), 0, 0)', '6'],
		];
		for (const [expr, error] of cases) {
			expect(errors(expr), expr).toEqual([error]);
		}
	});

	it('stays quiet where the result fits, or the date string is ambiguous', () => {
		for (const expr of [
			'DateAdd("d", -1, DateValue("12/31/9999"))',
			'DateSerial(2000, DateDiff("d", #1/1/2000#, #1/31/2000#), 1)',
			'DateDiff("s", #1/1/100#, #12/31/9999#)',
			'DateAdd("d", Day(#1/2/2000#), DateValue("12/29/9999"))',
			'DateAdd("d", 1, DateValue("1/2/9999"))',
			'TimeSerial(0, DateDiff("n", #1/1/2000#, #1/2/2000#), 0)',
		]) {
			expect(errors(expr), expr).toEqual([]);
		}
	});
});
