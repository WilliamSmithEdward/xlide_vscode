// Diagnostics tests: a sheet's Visible, reached late-bound through ActiveSheet,
// Worksheets(1) or Sheets(1), refuses a String that is no number with 1004
// (issue #416). Each case was run through pyVBAharness on 2026-10-02 in
// Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function messages(line: string): string[] {
	const src = `Option Explicit\nSub T()\n    ${line}\nEnd Sub\n`;
	return analyzeModule(src).filter((diag) => diag.code === 'host-property-value-out-of-range').map((diag) => diag.message);
}

describe("a sheet's Visible reached late-bound (issue #416)", () => {
	it('refuses a String that is no number, True included, with 1004', () => {
		for (const line of ['ActiveSheet.Visible = "abc"', 'Worksheets(1).Visible = "abc"', 'Sheets(1).Visible = "abc"', 'ActiveSheet.Visible = "True"']) {
			expect(messages(line), line).toEqual([expect.stringContaining("Run-time error '1004'")]);
		}
	});

	it('stays quiet on a number, True and the constants', () => {
		for (const line of ['ActiveSheet.Visible = True', 'Worksheets(1).Visible = xlSheetVisible', 'ActiveSheet.Visible = "2"']) {
			expect(messages(line), line).toEqual([]);
		}
	});
});
