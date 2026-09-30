// Diagnostics tests: ThisWorkbook.Sheets, Worksheets and Charts given a name
// or index the saved workbook lacks (issue #229). Measured through
// pyVBAharness in Excel 16.0 on 2026-09-30: a missing name, an index past the
// count, Charts(1) with no chart sheet and Worksheets("<a chart sheet>") raise
// 9; names match whatever their case. Sheets holds every kind of sheet,
// Worksheets only worksheets (not an Excel 4 macro sheet), Charts only chart
// sheets.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { mergeSheetChanges, sheetChangesIn, type WorkbookSheetInfo } from '../../src/analyzer/symbols/sheetChanges';
import { byCode } from '../helpers/diagnostics';

const CODE = 'sheet-not-in-workbook';

// SheetsFixture's sheets: Budget, Drawn, Trend (a chart sheet), Later, Hidden.
const SHEETS: WorkbookSheetInfo[] = [
	{ name: 'Budget', kind: 'worksheet' },
	{ name: 'Drawn', kind: 'worksheet' },
	{ name: 'Trend', kind: 'chartsheet' },
	{ name: 'Later', kind: 'worksheet' },
	{ name: 'Hidden', kind: 'worksheet' },
];

function wrap(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

/** The findings on `src`, analyzed with the project made of it and `others`. */
function hits(src: string, others: string[] = [], sheets: WorkbookSheetInfo[] | null = SHEETS) {
	return byCode(analyzeModule(src, {
		workbookSheets: sheets ?? undefined,
		projectSheetChanges: mergeSheetChanges([src, ...others].map(sheetChangesIn)),
	}), CODE);
}

describe('sheet-not-in-workbook (issue #229)', () => {
	it('flags a name the workbook lacks, and an index past its last sheet', () => {
		const cases: Array<[string, string]> = [
			['Set Main = ThisWorkbook.Sheets("doesntexist")', "this workbook has no sheet named 'doesntexist'"],
			['Set Main = ThisWorkbook.Sheets(10)', 'This workbook has 5 sheets, so index 10 is past the last'],
			['Main = ThisWorkbook.Worksheets(5).Name', 'This workbook has 4 worksheets, so index 5'],
			['Main = ThisWorkbook.Charts(2).Name', 'This workbook has 1 chart sheet, so index 2'],
			['Main = ThisWorkbook.Worksheets("Trend").Name', "'Trend' is a chart sheet, not a worksheet"],
			['Main = ThisWorkbook.Charts("Budget").Name', "'Budget' is a worksheet, not a chart sheet"],
			['Main = ThisWorkbook.Sheets.Item("Nope").Name', "no sheet named 'Nope'"],
			['Main = Application.ThisWorkbook.Worksheets("Nope").Name', "no worksheet named 'Nope'"],
			['Main = ThisWorkbook.Sheets("1").Name', "no sheet named '1'"],
		];
		for (const [line, text] of cases) {
			const found = hits(wrap(line));
			expect(found, line).toHaveLength(1);
			expect(found[0].message.toLowerCase(), line).toContain(text.toLowerCase());
			expect(found[0].message, line).toContain("Run-time error '9': Subscript out of range");
		}
	});

	it('marks the argument', () => {
		const src = wrap('Set Main = ThisWorkbook.Sheets("doesntexist")');
		const [hit] = hits(src);
		expect(src.slice(hit.span.start, hit.span.end)).toBe('"doesntexist"');
	});

	it('leaves names and indexes the workbook has, whatever their case', () => {
		for (const line of [
			'Main = ThisWorkbook.Sheets("BUDGET").Name',
			'Main = ThisWorkbook.Sheets("hidden").Name',
			'Main = ThisWorkbook.Sheets(5).Name',
			'Main = ThisWorkbook.Worksheets(4).Name',
			'Main = ThisWorkbook.Charts(1).Name',
			'Main = ThisWorkbook.Charts("trend").Name',
			// Rounds to 1, and runs.
			'Main = ThisWorkbook.Sheets(1.4).Name',
			'Main = ThisWorkbook.Sheets.Count',
		]) {
			expect(hits(wrap(line)), line).toEqual([]);
		}
	});

	it('leaves a workbook other than ThisWorkbook alone: the active one may be any', () => {
		for (const line of [
			'Main = Sheets("Nope").Name',
			'Main = Worksheets(9).Name',
			'Main = ActiveWorkbook.Sheets("Nope").Name',
			'Main = Workbooks(1).Sheets("Nope").Name',
			'Main = x.ThisWorkbook.Sheets("Nope").Name',
		]) {
			expect(hits(wrap(line)), line).toEqual([]);
		}
	});

	it('leaves index 0 to host-argument-out-of-range', () => {
		expect(hits(wrap('Main = ThisWorkbook.Sheets(0).Name'))).toEqual([]);
	});

	it('flags the pass of a loop that runs past the last sheet', () => {
		const src = wrap('Dim i As Long', 'For i = 1 To 6', '    Main = ThisWorkbook.Sheets(i).Name', 'Next');
		expect(hits(src)).toHaveLength(1);
		expect(hits(wrap('Dim i As Long', 'For i = 1 To 5', '    Main = ThisWorkbook.Sheets(i).Name', 'Next'))).toEqual([]);
	});

	it('checks nothing without the saved sheets or the project scan', () => {
		const src = wrap('Set Main = ThisWorkbook.Sheets("doesntexist")');
		expect(byCode(analyzeModule(src), CODE)).toEqual([]);
		expect(byCode(analyzeModule(src, { workbookSheets: SHEETS }), CODE)).toEqual([]);
		expect(hits(src, [], null)).toEqual([]);
	});

	it('checks nothing when a local is named ThisWorkbook, or the name is outside ASCII', () => {
		expect(hits(wrap('Dim ThisWorkbook As Object', 'Main = ThisWorkbook.Sheets("Nope").Name'))).toEqual([]);
		expect(hits(wrap(`Main = ThisWorkbook.Sheets("${String.fromCharCode(0xc9)}t${String.fromCharCode(0xe9)}").Name`))).toEqual([]);
	});

	describe('a sheet the code itself may make', () => {
		const MISSING = wrap('Main = ThisWorkbook.Sheets("Report").Name', 'Main = ThisWorkbook.Sheets(7).Name');

		it('leaves the index and the default names alone once any module adds a sheet', () => {
			for (const adder of [
				'Sub A()\n    ThisWorkbook.Worksheets.Add\nEnd Sub\n',
				'Sub A()\n    Sheets.Add After:=Sheets(1)\nEnd Sub\n',
				'Sub A()\n    Charts.Add2\nEnd Sub\n',
				'Sub A()\n    With ThisWorkbook.Worksheets\n        .Add\n    End With\nEnd Sub\n',
				'Sub A()\n    Sheet1.Copy After:=Sheet1\nEnd Sub\n',
			]) {
				expect(hits(wrap('Main = ThisWorkbook.Sheets(7).Name'), [adder]), adder).toEqual([]);
				expect(hits(wrap('Main = ThisWorkbook.Sheets("Sheet6").Name'), [adder]), adder).toEqual([]);
				expect(hits(wrap('Main = ThisWorkbook.Sheets("Budget (2)").Name'), [adder]), adder).toEqual([]);
				// A name no added sheet is given is still missing.
				expect(hits(wrap('Main = ThisWorkbook.Sheets("Report").Name'), [adder]), adder).toHaveLength(1);
			}
		});

		it('does not take a range copy or a collection add for a new sheet', () => {
			for (const other of [
				'Sub A()\n    Range("A1").Copy Range("B1")\nEnd Sub\n',
				'Sub A()\n    Dim c As New Collection\n    c.Add 1\nEnd Sub\n',
				'Sub A()\n    Sheet1.Copy\nEnd Sub\n',
			]) {
				expect(hits(MISSING, [other]), other).toHaveLength(2);
			}
		});

		it('leaves a name some code gives a sheet', () => {
			expect(hits(MISSING, ['Sub A()\n    Worksheets.Add.Name = "Report"\nEnd Sub\n'])).toEqual([]);
			expect(hits(wrap('Main = ThisWorkbook.Sheets("Report").Name'), ['Sub A()\n    If x Then Sheet1.Name = "report"\nEnd Sub\n'])).toEqual([]);
		});

		it('leaves every name once code gives a sheet a computed name, but not for a comparison', () => {
			for (const renamer of [
				'Sub A(s As String)\n    Sheet1.Name = s\nEnd Sub\n',
				'Sub A(s As String)\n    Me.Name = "R" & s\nEnd Sub\n',
				'Sub A(s As String)\n    With Sheet1\n        .Name = s\n    End With\nEnd Sub\n',
				'Sub A(s As String)\n    ThisWorkbook.Worksheets(1).Name = s\nEnd Sub\n',
			]) {
				expect(hits(wrap('Main = ThisWorkbook.Sheets("Report").Name'), [renamer]), renamer).toEqual([]);
			}
			expect(hits(wrap('Main = ThisWorkbook.Sheets("Report").Name'), ['Sub A(s As String)\n    If Sheet1.Name = s Then Exit Sub\nEnd Sub\n'])).toHaveLength(1);
		});

		it('reads nothing from comments', () => {
			expect(hits(MISSING, ["Sub A()\n    ' Worksheets.Add\n    Rem Sheet1.Name = s\nEnd Sub\n"])).toHaveLength(2);
		});
	});
});

describe('sheet-not-in-workbook against the measured workbook (issue #229)', () => {
	// The harness workbook has one sheet, Sheet1. These are the oracle's
	// issue229 cases, each with what Excel did.
	const HARNESS: WorkbookSheetInfo[] = [{ name: 'Sheet1', kind: 'worksheet' }];
	it.each([
		['Set Main = ThisWorkbook.Sheets("doesntexist")', 1],
		['Set Main = ThisWorkbook.Sheets(10)', 1],
		['Main = ThisWorkbook.Worksheets("SHEET1").Name', 0],
		['Main = ThisWorkbook.Charts(1).Name', 1],
		['Main = ThisWorkbook.Sheets(" Sheet1").Name', 1],
		['Main = ThisWorkbook.Sheets(Array("Sheet1")).Count', 0],
		['Main = TypeName(ThisWorkbook.Sheets.Item("nope"))', 1],
		['Main = ThisWorkbook.Sheets(1.4).Name', 0],
		['Main = ThisWorkbook.Sheets("1").Name', 1],
	] as const)('%s', (line, count) => {
		expect(hits(wrap(line), [], HARNESS)).toHaveLength(count);
	});
});

describe('sheetChangesIn (issue #229)', () => {
	it('names what a module does to sheets', () => {
		const changes = sheetChangesIn('Sub A(s)\n    Worksheets.Add.Name = "Out"\n    If s = 1 Then Sheet2.Name = "Two" Else Sheet3.Name = "Three"\nEnd Sub\n');
		expect(changes.addsSheets).toBe(true);
		expect([...changes.namesAssigned].sort()).toEqual(['out', 'three', 'two']);
		expect(changes.assignsComputedName).toBe(false);
	});
});
