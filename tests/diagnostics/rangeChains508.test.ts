// Diagnostics tests: chains of Range members off the sheet (issue #508).
// Each sample was measured through pyVBAharness on 2026-10-02 in Excel 16.0
// (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode } from '../helpers/diagnostics';

const RANGE = 'host-argument-out-of-range';

function wrap(expr: string): string {
	return `Option Explicit\nFunction Main() As Variant\n    Main = ActiveSheet.${expr}.Address\nEnd Function\n`;
}

describe('chains of Range members (issue #508)', () => {
	it('reports the step that leaves the sheet', () => {
		const cases: Array<[string, string]> = [
			['Range("A1").Offset(2).Offset(-4, 0)', 'row -1'],
			['Range("B2:C2").Offset(0, -2).Offset(0, -3)', 'columns 0 to 1'],
			['Range("A1048576").Offset(0, 4).Offset(2, 0)', 'row 1048578'],
			['Range("A1:A3").Offset(-2)', 'rows -1 to 1'],
			['Range("XFC1048575:XFD1048576").Resize(6)', 'rows 1048575 to 1048580'],
			['Range("XFC1048575:XFD1048576").Resize(3, 4)', 'columns 16383 to 16386'],
			['Range("XFC1048575:XFD1048576").Cells(3, 5)', 'row 1048577'],
			['Range("XFC1048575:XFD1048576").Item(5)', 'row 1048577'],
			['Range("A1048576").Item(6)', 'row 1048581'],
			['Range("A1").EntireRow.Offset(1, 5)', 'columns 6 to 16389'],
			['Range("A1").Columns(2).EntireColumn.Offset(1)', 'rows 2 to 1048577'],
			['Range("XFD1").EntireColumn.Resize(2, 3)', 'columns 16384 to 16386'],
			['Range("A1").EntireColumn.Columns(-1)', 'column -1'],
		];
		for (const [expr, where] of cases) {
			const hits = byCode(analyzeModule(wrap(expr)), RANGE);
			expect(hits, expr).toHaveLength(1);
			expect(hits[0].message, expr).toContain(where);
		}
	});

	it('stays quiet while the block stays on the sheet', () => {
		for (const expr of [
			'Range("A1").Offset(2).Offset(-2, 0)',
			'Range("A2:B3").Offset(-1, 0)',
			'Range("A1").EntireRow.Offset(1, 0)',
			'Range("A1").EntireColumn.Offset(0, 1)',
			'Range("XFC1048575:XFD1048576").Item(4)',
			'Range("A1:B2").Resize(3).Offset(1, 1)',
			'Range("B2").Cells(0, 0)',
		]) {
			expect(byCode(analyzeModule(wrap(expr)), RANGE), expr).toHaveLength(0);
		}
	});
});
