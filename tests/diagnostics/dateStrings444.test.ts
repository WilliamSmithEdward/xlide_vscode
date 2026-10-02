// Diagnostics tests: date strings no locale reads (issue #444). Measured in
// Excel 16.0 (build 20326, en-US, 2026-10-02), each function given each
// string as a literal.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function reported(call: string): boolean {
	const src = `Option Explicit\nFunction Main() As Variant\n    Main = ${call}\nEnd Function\n`;
	return analyzeModule(src).some((diag) => diag.code === 'runtime-conversion-value' || diag.code === 'runtime-argument-value');
}

const READERS = ['DateValue', 'TimeValue', 'CDate', 'Year', 'Month', 'Day', 'Hour', 'Minute', 'Weekday'];

describe('a date string no locale reads', () => {
	it('raises 13 for a year past 9999, in every reader', () => {
		for (const fn of READERS) {
			expect(reported(`${fn}("1/1/10000")`), fn).toBe(true);
		}
	});

	it('raises 13 for an hour of 25, in every reader', () => {
		for (const fn of READERS) {
			expect(reported(`${fn}("25:00")`), fn).toBe(true);
		}
	});

	it('raises 13 for a whole number in DateValue and TimeValue, not in CDate', () => {
		for (const call of ['DateValue("12")', 'DateValue("-1")', 'TimeValue("-1")', 'TimeValue("12")']) {
			expect(reported(call), call).toBe(true);
		}
		for (const call of ['CDate("12")', 'Year("12")', 'CDate("-1")']) {
			expect(reported(call), call).toBe(false);
		}
	});

	it('runs for a date the readers take, and IsDate never raises', () => {
		for (const fn of READERS) {
			for (const s of ['1/1/2000', '12/31/9999']) {
				expect(reported(`${fn}("${s}")`), `${fn}("${s}")`).toBe(false);
			}
		}
		for (const s of ['1/1/10000', '25:00', '12', '-1', 'abc', '']) {
			expect(reported(`IsDate("${s}")`), s).toBe(false);
		}
	});
});
