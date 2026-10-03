// Diagnostics tests: a sheet protected with no password or with Allow flags,
// and a workbook whose structure is protected (issue #684). Each sample was
// measured through pyVBAharness on 2026-10-03 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode } from '../helpers/diagnostics';

const RANGE = 'host-argument-out-of-range';

function hits(setup: readonly string[], ...lines: string[]) {
	const src = `Option Explicit\nFunction Main() As Variant\n${[...setup, ...lines].map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
	return byCode(analyzeModule(src), RANGE);
}

const SHEET = ['Dim ws As Worksheet', 'Set ws = Worksheets.Add'];
const BOOK = ['Dim wb As Workbook, ws As Worksheet', 'Set wb = Workbooks.Add', 'Set ws = wb.Worksheets.Add'];

describe('a sheet protected with no password, or with Allow flags (issue #684)', () => {
	it('stays quiet where Excel runs the statement', () => {
		for (const lines of [
			['ws.Protect', 'ws.Unprotect "nope"'],
			['ws.Protect ""', 'ws.Unprotect "nope"'],
			['ws.Protect AllowInsertingRows:=True', 'ws.Rows(1).Insert'],
			['ws.Protect AllowInsertingRows:=True', 'ws.Range("A1").EntireRow.Insert'],
			['ws.Protect "pw", AllowInsertingRows:=True', 'ws.Rows(1).Insert'],
			['ws.Protect AllowInsertingColumns:=True', 'ws.Columns(1).Insert'],
			['ws.Protect AllowFormattingCells:=True', 'ws.Range("A1").Font.Bold = True'],
			['ws.Protect AllowFormattingColumns:=True', 'ws.Columns(1).ColumnWidth = 20'],
		]) {
			expect(hits(SHEET, ...lines), lines.join(' / ')).toEqual([]);
		}
	});

	it('still reports what the flags do not allow', () => {
		for (const lines of [
			['ws.Protect AllowDeletingRows:=True', 'ws.Rows(1).Delete'],
			['ws.Protect AllowDeletingColumns:=True', 'ws.Columns(1).Delete'],
			['ws.Protect AllowInsertingRows:=True', 'ws.Range("A1").Value = 1'],
			['ws.Protect AllowInsertingRows:=True', 'ws.Columns(1).Insert'],
			['ws.Protect AllowFormattingCells:=True', 'ws.Range("A1").ClearContents'],
			['ws.Protect AllowFormattingCells:=True', 'ws.Range("A1").Value = 1'],
			['ws.Protect AllowInsertingRows:=False', 'ws.Rows(1).Insert'],
			['ws.Protect "pw"', 'ws.Unprotect "nope"'],
		]) {
			const found = hits(SHEET, ...lines);
			expect(found, lines.join(' / ')).toHaveLength(1);
			expect(found[0].message).toContain("'1004'");
		}
	});
});

describe("a workbook whose structure is protected (issue #684)", () => {
	it.each([
		['adding a sheet', ['wb.Protect "pw", Structure:=True', 'wb.Worksheets.Add'], "'wb' has its structure protected here, so no sheet can be added. This will raise Run-time error '1004': Method 'Add' of object 'Sheets' failed."],
		['adding one to the active workbook', ['ActiveWorkbook.Protect "pw", Structure:=True', 'Worksheets.Add'], "The active workbook has its structure protected here"],
		['adding one through Sheets', ['wb.Protect "pw", True', 'wb.Sheets.Add'], "Method 'Add' of object 'Sheets' failed."],
		['Protect with Structure left out', ['wb.Protect "pw"', 'wb.Worksheets.Add'], "Method 'Add' of object 'Sheets' failed."],
		['Protect with no password', ['wb.Protect', 'wb.Worksheets.Add'], "Method 'Add' of object 'Sheets' failed."],
		['a rename', ['wb.Protect "pw", True', 'ws.Name = "zqRenamed"'], "Method 'Name' of object '_Worksheet' failed."],
		['a Delete', ['wb.Protect "pw", True', 'Application.DisplayAlerts = False', 'ws.Delete'], "Method 'Delete' of object '_Worksheet' failed."],
		['hiding a sheet', ['wb.Protect "pw", True', 'ws.Visible = xlSheetHidden'], "Method 'Visible' of object '_Worksheet' failed."],
		['a Copy', ['wb.Protect "pw", True', 'ws.Copy After:=ws'], "Workbook is protected and cannot be changed."],
		['Unprotect with another password', ['wb.Protect "pw", True', 'wb.Unprotect "nope"'], "The password you supplied is not correct."],
	])('reports %s', (_label, lines, message) => {
		const found = hits(BOOK, ...lines);
		expect(found, lines.join(' / ')).toHaveLength(1);
		expect(found[0].message).toContain(message);
	});

	it('stays quiet once unprotected, with Structure False, and on reads and cell writes', () => {
		for (const lines of [
			['wb.Protect "pw", Structure:=True', 'wb.Unprotect "pw"', 'wb.Worksheets.Add'],
			['wb.Protect Structure:=False, Windows:=True', 'wb.Worksheets.Add'],
			['wb.Protect "pw", True', 'Main = ws.Name'],
			['wb.Protect "pw", True', 'ws.Range("A1").Value = 5'],
			['ActiveWorkbook.Protect "pw", True', 'Workbooks.Add', 'Worksheets.Add'],
		]) {
			expect(hits(BOOK, ...lines), lines.join(' / ')).toEqual([]);
		}
	});
});
