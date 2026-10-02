// Diagnostics tests: host arguments held in a local (issue #345). Each
// raising sample was measured through pyVBAharness on 2026-10-02 in Excel
// 16.0 (build 20326); each quiet one runs there.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode } from '../helpers/diagnostics';

const RANGE = 'host-argument-out-of-range';

function wrap(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

describe('host-argument-out-of-range from a local (issue #345)', () => {
	it('reads a number local whatever its type', () => {
		for (const type of ['Long', 'Integer', 'Double', 'Variant']) {
			const src = wrap(`Dim x As ${type}`, 'x = 0', 'Main = Cells(x, 1).Row');
			const hits = byCode(analyzeModule(src), RANGE);
			expect(hits, type).toHaveLength(1);
			expect(hits[0].message, type).toContain("'1004'");
		}
	});

	it('reads a local in Cells, Rows, Columns, Offset, Resize and Worksheets', () => {
		const cases: Array<[string, string, string]> = [
			['0', 'Main = Cells(1, x).Row', "'1004'"],
			['0', 'Main = Rows(x).Row', "'1004'"],
			['0', 'Main = Columns(x).Column', "'1004'"],
			['1048577', 'Main = Cells(x, 1).Row', "'1004'"],
			['-1', 'Main = Range("A1").Offset(x).Row', "'1004'"],
			['0', 'Main = Range("A1").Resize(x).Row', "'1004'"],
			['0', 'Main = Worksheets(x).Name', "'9'"],
		];
		for (const [value, line, error] of cases) {
			const src = wrap('Dim x As Long', `x = ${value}`, line);
			const hits = byCode(analyzeModule(src), RANGE);
			expect(hits, line).toHaveLength(1);
			expect(hits[0].message, line).toContain(error);
		}
	});

	it('reads a Long nothing assigns as 0', () => {
		const src = wrap('Dim x As Long', 'Main = Cells(x, 1).Row');
		expect(byCode(analyzeModule(src), RANGE)).toHaveLength(1);
	});

	it('reads a String local as an address or a column', () => {
		const cases: Array<[string, string, string]> = [
			['A1:', 'Main = Range(s).Row', "'1004'"],
			['AAAA', 'Main = Cells(1, s).Row', "'13'"],
		];
		for (const [value, line, error] of cases) {
			const src = wrap('Dim s As String', `s = "${value}"`, line);
			const hits = byCode(analyzeModule(src), RANGE);
			expect(hits, line).toHaveLength(1);
			expect(hits[0].message, line).toContain(error);
		}
	});

	it('stays quiet for a good value, a reassigned local and an address that may be a name', () => {
		const lines = [
			['Dim s As String', 's = "B2"', 'Main = Range(s).Row'],
			['Dim r As Long', 'r = 1', 'Main = Cells(r, 1).Row'],
			['Dim x As Long', 'x = 0', 'x = 2', 'Main = Cells(x, 1).Row'],
			// A0 can be a defined name, and Range("A0") then runs.
			['Dim s As String', 's = "A0"', 'Main = Range(s).Row'],
		];
		for (const body of lines) {
			expect(byCode(analyzeModule(wrap(...body)), RANGE), body.join(' / ')).toHaveLength(0);
		}
	});
});
