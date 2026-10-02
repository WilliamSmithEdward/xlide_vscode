// Diagnostics tests: WorksheetFunction names it lacks (438) and literals a
// worksheet function refuses (1004) (issue #442). Measured in Excel 16.0
// 64-bit (build 20430, 2026-10-02) through pyVBAharness.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(expr: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    Main = ${expr}\nEnd Function\n`;
	return analyzeModule(src, { host: 'excel' }).filter((diag) => diag.severity === 'error').map((diag) => `${diag.code}: ${diag.message}`);
}

describe('a name WorksheetFunction does not have', () => {
	it('raises 438', () => {
		for (const expr of [
			'WorksheetFunction.Mid("abc", 1, 1)', 'WorksheetFunction.Left("abc", 1)', 'WorksheetFunction.Summ(1, 2)',
			'Application.WorksheetFunction.Summ(1, 2)', 'WorksheetFunction.Sqrt(4)', 'WorksheetFunction.Ifs(True, 1)',
		]) {
			expect(errors(expr), expr).toEqual([expect.stringMatching(/^runtime-member-not-found: .*'438'/)]);
		}
	});

	it('is quiet for a worksheet function', () => {
		for (const expr of ['WorksheetFunction.Sum(1, 2)', 'Application.WorksheetFunction.Max(1, 2)', 'WorksheetFunction.XLookup(2, Array(1, 2), Array("a", "b"))']) {
			expect(errors(expr), expr).toEqual([]);
		}
	});
});

describe('literals a worksheet function refuses', () => {
	it('raise 1004 through WorksheetFunction', () => {
		for (const expr of [
			'WorksheetFunction.Ln(0)', 'WorksheetFunction.Ln(-1)', 'WorksheetFunction.Log10(0)', 'WorksheetFunction.Power(-1, 0.5)',
			'WorksheetFunction.Power(0, -1)', 'WorksheetFunction.Sum("abc")', 'WorksheetFunction.Max("abc")', 'WorksheetFunction.Average("x", 1)',
			'WorksheetFunction.Dec2Bin(1000)', 'WorksheetFunction.Dec2Bin(-513)', 'WorksheetFunction.Large(Array(1, 2), 3)',
			'WorksheetFunction.Small(Array(1, 2), 0)', 'WorksheetFunction.Index(Array(1, 2), 3)', 'WorksheetFunction.Index(Array(1, 2), 5)',
			'WorksheetFunction.Match("zzz", Array("a", "b"), 0)', 'WorksheetFunction.Match(3, Array(1, 2), 0)',
		]) {
			expect(errors(expr), expr).toEqual([expect.stringMatching(/^host-argument-out-of-range: .*'1004'/)]);
		}
	});

	it('run otherwise, and through Application', () => {
		for (const expr of [
			'WorksheetFunction.Ln(1)', 'WorksheetFunction.Power(-8, 2)', 'WorksheetFunction.Power(2, 0.5)', 'WorksheetFunction.Sum("5")',
			'WorksheetFunction.Dec2Bin(511)', 'WorksheetFunction.Dec2Bin(-512)', 'WorksheetFunction.Large(Array(1, 2), 2)',
			'WorksheetFunction.Index(Array(1, 2), 2)', 'WorksheetFunction.Match("B", Array("a", "b"), 0)', 'WorksheetFunction.Match(2, Array(1, 2), 0)',
			'IsError(Application.Ln(0))', 'IsError(Application.Match("zzz", Array("a", "b"), 0))',
		]) {
			expect(errors(expr), expr).toEqual([]);
		}
	});
});
