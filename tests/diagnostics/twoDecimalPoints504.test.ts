// Diagnostics tests: a string with a second decimal point under either
// convention, "." or "," (issue #504). Measured on 2026-10-03 in Excel 16.0
// (build 20430, en-US).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(call: string): string[] {
	return analyzeModule(`Option Explicit\nFunction Main() As Variant\n    Main = ${call}\nEnd Function\n`)
		.filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('two decimal points under either convention (issue #504)', () => {
	it('reports a string with two "." and two ","', () => {
		for (const call of ['CDbl("1,2.3,4.5")', 'CLng("1.2,3.4,5")', 'CDbl("1,2.3.4,5")', 'CInt("1..2,,3")', 'CDbl(" 1,2.3,4.5 ")', 'CDbl("$1,2.3,4.5")']) {
			expect(errors(call), call).toEqual(['runtime-conversion-value']);
		}
	});

	it('leaves alone a string some locale reads as a number', () => {
		// "1.5.5" and "1.2,3.4" raise 13 under en-US, and read as 155 and
		// 12.34 where "." groups thousands.
		for (const call of ['CDbl("1.5.5")', 'CDbl("1.2,3.4")', 'CDbl("1,5,5")', 'CDbl("1,2.3")', 'CLng("1,2,3.4")']) {
			expect(errors(call), call).toEqual([]);
		}
	});
});
