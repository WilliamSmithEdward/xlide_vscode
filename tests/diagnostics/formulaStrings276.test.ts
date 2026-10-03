// Diagnostics tests: a formula String Excel cannot parse, a warning since a
// cell formatted as Text takes it (issue #276). Measured on 2026-10-03 in
// Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function warnings(body: string): string[] {
	return analyzeModule(`Option Explicit\nFunction Main() As Variant\n    ${body}\n    Main = 1\nEnd Function\n`)
		.filter((diag) => diag.code === 'formula-string-unparsed').map((diag) => diag.severity);
}

describe('a formula String Excel cannot parse (issue #276)', () => {
	it('warns on a parenthesis or string left open, or an operator last', () => {
		for (const body of [
			'ActiveSheet.Range("A1").Formula = "=SUM(B1:B2"',
			'ActiveSheet.Range("A1").Formula = "=""abc"',
			'ActiveSheet.Range("A1").FormulaR1C1 = "=SUM(R1C2:R2C2"',
			'ActiveSheet.Range("A1").Value = "=SUM(B1"',
			'ActiveSheet.Range("A1").Formula = "=1+"',
			'ActiveSheet.Range("A1").Formula = "=SUM(B1))"',
			'ActiveSheet.Range("A1").Formula2 = "=A1&"',
			'ActiveSheet.Range("A1").Value2 = "=(1"',
		]) {
			expect(warnings(body), body).toEqual(['warning']);
		}
	});

	it('stays quiet on a formula Excel takes, and on plain text', () => {
		for (const body of [
			'ActiveSheet.Range("A1").Formula = "=SUM(B1:B2)"',
			'ActiveSheet.Range("A1").Formula = "=""("""',
			'ActiveSheet.Range("A1").Formula = "=NOSUCHFN(1)"',
			'ActiveSheet.Range("A1").Formula = "abc("',
			'ActiveSheet.Range("A1").Formula = "=IF(B1="""",1,2)"',
			'ActiveSheet.Range("A1").Formula = "=-1"',
			'ActiveSheet.Range("A1").Value = "=1%"',
		]) {
			expect(warnings(body), body).toEqual([]);
		}
	});
});
