// Diagnostics tests: Excel errors literal ranges and a sheet the code just
// added prove (issue #472). Each sample was measured through pyVBAharness on
// 2026-10-02 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode } from '../helpers/diagnostics';

const RANGE = 'host-argument-out-of-range';
const SETUP = ['Dim w1 As Worksheet, w2 As Worksheet', 'Set w1 = ActiveSheet', 'Set w2 = Worksheets.Add'];

function hits(...lines: string[]) {
	const src = `Option Explicit\nFunction Main() As Variant\n${[...SETUP, ...lines].map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
	return byCode(analyzeModule(src), RANGE);
}

describe('errors literal ranges and a new sheet prove (issue #472)', () => {
	it('does not assume a sheet loaded from a template is empty', () => {
		expect(hits('Set w2 = Worksheets.Add(Type:="Report.xltx")', 'Main = w2.Cells.Find("zzz").Row')).toEqual([]);
	});

	it('does not assume a handled Add succeeded', () => {
		expect(hits('On Error Resume Next', 'Set w2 = Worksheets.Add', 'On Error GoTo 0', 'Main = w2.Cells.Find("zzz").Row')).toEqual([]);
	});

	it.each([
		'ActiveSheet.Range("A1").Value = "zzz"',
		'Worksheets(1).Range("A1").Value = "zzz"',
		'Dim alias As Worksheet\n    Set alias = w2\n    alias.Range("A1").Value = "zzz"',
		'Application.Run "PopulateSheet"',
		'DoEvents',
	])('does not assume a new sheet stays empty after another reference or call: %s', change => {
		expect(hits(change, 'Main = w2.Cells.Find("zzz").Row')).toEqual([]);
		expect(hits(change, 'Main = w2.Cells.SpecialCells(xlCellTypeConstants).Count')).toEqual([]);
	});

	it('reports literal range failures without inferring new-sheet contents', () => {
		const cases: Array<[string, string]> = [
			['Main = Intersect(w2.Range("A1"), w2.Range("B2")).Count', "'91'"],
			['Main = Union(w1.Range("A1"), w2.Range("A1")).Count', "'1004'"],
			['w2.ShowAllData', "'1004'"],
			['Main = w2.Cells.SpecialCells(xlCellTypeConstants).Count', "'1004'"],
			['Main = w2.UsedRange.SpecialCells(xlCellTypeFormulas).Count', "'1004'"],
			['Main = w2.Range("A1:A3").SpecialCells(xlCellTypeBlanks).Count', "'1004'"],
			['Main = w2.Cells.Find("zzz").Row', "'91'"],
			['w2.Range("A1:B5").AutoFilter Field:=1, Criteria1:="x"', "'1004'"],
			['w2.Range("A1:A3").TextToColumns Destination:=w2.Range("B1")', "'1004'"],
		];
		for (const [line, error] of cases) {
			const found = hits(line);
			expect(found, line).toHaveLength(/Intersect|Union/.test(line) ? 1 : 0);
			if (found.length) { expect(found[0].message, line).toContain(error); }
		}
	});

	it('stays quiet where the ranges meet, the sheet is one, or the code wrote to it', () => {
		for (const lines of [
			['Main = Intersect(w2.Range("A1:B2"), w2.Range("B2:C3")).Count'],
			['Main = Union(w2.Range("A1"), w2.Range("C3")).Count'],
			['Main = w2.Cells.Find("zzz") Is Nothing'],
			['Main = w2.AutoFilterMode'],
			['w2.Range("A1").Value = 5', 'Main = w2.Cells.SpecialCells(xlCellTypeConstants).Count'],
			['w2.Range("A1").Value = "zzz"', 'Main = w2.Cells.Find("zzz").Row'],
		]) {
			expect(hits(...lines), lines.join(' / ')).toHaveLength(0);
		}
	});
});
