// Diagnostics tests: addresses and columns a literal decides (issue #276).
// Measured in Excel 16.0 64-bit (2026-10-02) through pyVBAharness.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function found(body: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    Dim ws As Worksheet\n    Set ws = ActiveSheet\n    ${body}\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('an address or a column no sheet has', () => {
	it('refuses an R1C1-style address, which no name may be', () => {
		expect(found('Main = ws.Range("R1C1").Address')).toEqual(['host-argument-out-of-range']);
	});

	it('refuses a column past XFD', () => {
		expect(found('Main = ws.Columns("XFE").Address')).toEqual(['host-argument-out-of-range']);
		expect(found('Main = ws.Columns("XFD").Address')).toEqual([]);
	});

	it('leaves alone what a workbook name may stand for', () => {
		for (const address of ['A0', 'XFE1', 'A1048577', 'A1:ZZZZ1']) {
			expect(found(`Main = ws.Range("${address}").Address`), address).toEqual([]);
		}
	});
});
