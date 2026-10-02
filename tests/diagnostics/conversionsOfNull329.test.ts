// Diagnostics tests: a conversion function given Null raises 94 (issue
// #329). Measured in Excel 16.0 64-bit (build 20430, 2026-10-02) through
// pyVBAharness.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(expr: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    Main = ${expr}\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => `${diag.code}: ${diag.message}`);
}

describe('a conversion of Null', () => {
	it('raises 94', () => {
		for (const fn of ['CBool', 'CByte', 'CCur', 'CDate', 'CDbl', 'CDec', 'CInt', 'CLng', 'CLngLng', 'CSng', 'CStr']) {
			expect(errors(`${fn}(Null)`), fn).toEqual([expect.stringContaining("'94'")]);
		}
	});

	it('is quiet for CVar, which keeps the Null', () => {
		expect(errors('CVar(Null)')).toEqual([]);
	});
});
