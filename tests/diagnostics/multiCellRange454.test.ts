// Diagnostics tests: a multi-cell Range read as one value, in the shapes
// multi-cell-range-as-scalar missed (issue #454). Measured in Excel 16.0
// 64-bit (build 20430, 2026-10-02) through pyVBAharness.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const CODE = 'multi-cell-range-as-scalar';

function found(...lines: string[]): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
	return analyzeModule(src, { host: 'excel' }).filter((diag) => diag.code === CODE).map((diag) => diag.message);
}

describe('a multi-cell Range read as one value', () => {
	it('raises 13 as a whole argument of a built-in that takes one value', () => {
		for (const call of [
			'CStr(Range("A1:A2").Value)', 'Len(Range("A1:A2"))', 'LenB(Range("A1:A2"))', 'Val(Range("A1:A2"))', 'CLng(Range("A1:A2"))',
			'CDbl(Range("A1:A2"))', 'CBool(Range("A1:A2"))', 'Trim(Range("A1:A2"))', 'UCase(Range("A1:A2"))', 'Left(Range("A1:A2"), 1)',
			'InStr(Range("A1:A2"), "a")', 'Abs(Range("A1:A2"))', 'Int(Range("A1:A2"))', 'Format(Range("A1:A2"))', 'Hex(Range("A1:A2"))',
		]) {
			expect(found(`Main = ${call}`), call).toEqual([expect.stringContaining("'13'")]);
		}
	});

	it('raises 13 as a Select Case subject, a Do While condition, or a two-cell, row or column Range', () => {
		expect(found('Select Case Range("A1:A2")', 'Case 1', 'End Select')).toHaveLength(1);
		expect(found('Select Case Range("A1:A2").Value', 'Case 1', 'End Select')).toHaveLength(1);
		expect(found('Do While Range("A1:A2") = 1', 'Loop')).toHaveLength(1);
		for (const value of ['Range("A1", "B2") + 1', 'Range("A1", "A2") & "x"', 'Rows(1) + 1', 'Columns(2) + 1']) {
			expect(found(`Main = ${value}`), value).toEqual([expect.stringContaining('two-dimensional array')]);
		}
	});

	it('is quiet where the array is taken whole, or the Range is one cell', () => {
		for (const line of [
			'Main = IsEmpty(Range("A1:A2"))', 'Main = IsNumeric(Range("A1:A2"))', 'Main = IsNull(Range("A1:A2").Text)', 'Main = Range("A1:A2")',
			'Main = WorksheetFunction.Sum(Range("A1:A2"))', 'Main = Range("A1") + 1', 'Main = Cells(1, 1) + 1', 'Main = TypeName(Range("A1:A2").Value)',
			'Main = UBound(Range("A1:A2").Value)', 'Main = IsArray(Range("A1:A2"))', 'Main = VarType(Range("A1:A2"))', 'Main = Range("A1", "A1") + 1',
			'Main = Len(Range("A1"))', 'Main = Rows(1).Row + 1',
		]) {
			expect(found(line), line).toEqual([]);
		}
		expect(found('Select Case Range("A1")', 'Case 1', 'End Select')).toEqual([]);
		expect(found('With Range("A1:A2")', 'Main = .Count', 'End With')).toEqual([]);
	});
});
