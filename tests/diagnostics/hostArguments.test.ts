// Diagnostics tests: host object model arguments the code proves wrong
// (issue #122). Each raising sample was measured through pyVBAharness on
// 2026-09-26 in Excel, Word or PowerPoint 16.0 (build 20326); each quiet one
// runs there.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { getWordObjectModel } from '../../src/analyzer/host/wordObjectModel';
import { getPowerPointObjectModel } from '../../src/analyzer/host/powerpointObjectModel';
import { byCode, expectDiagnostic, expectDiagnostics } from '../helpers/diagnostics';

const RANGE = 'host-argument-out-of-range';
const NAME = 'sheet-name-invalid';
const SCALAR = 'multi-cell-range-as-scalar';

function wrap(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

describe('host-argument-out-of-range - Excel (issue #122)', () => {
	it('flags index 0 into a 1-based collection, bare, qualified and through Item', () => {
		const cases: Array<[string, string]> = [
			['Main = Worksheets(0).Name', "'9'"],
			['Main = Sheets(-1).Name', "'9'"],
			['Main = Workbooks(0).Name', "'9'"],
			['Main = Names(0).Name', "'9'"],
			['Main = Worksheets.Item(0).Name', "'9'"],
			['Main = ThisWorkbook.Worksheets(0).Name', "'9'"],
			['Main = Rows(0).Row', "'1004'"],
			['Main = Columns(0).Column', "'1004'"],
			['Main = Range("A1").Areas(0).Address', "'1004'"],
			['Main = Worksheets(1).Shapes(0).Name', "'-2147024809'"],
		];
		for (const [line, error] of cases) {
			const src = wrap(line);
			const hits = byCode(analyzeModule(src), RANGE);
			expect(hits, line).toHaveLength(1);
			expect(hits[0].message, line).toContain(error);
		}
	});

	it('flags a row or column below 1 in Cells, Offset and Resize', () => {
		const cases: Array<[string, string]> = [
			['Main = Cells(0, 1).Value', '0'],
			['Main = Cells(1, 0).Value', '0'],
			['Main = Cells(0).Value', '0'],
			['Main = Worksheets(1).Cells(1, -1).Value', '-1'],
			['Main = Range("A1").Offset(-1, 0).Address', '-1, 0'],
			['Main = Range("A1").Offset(0, -1).Address', '0, -1'],
			['Main = Range("A1").Resize(0, 1).Address', '0'],
			['Main = Range("A1").Resize(1, 0).Address', '0'],
			['Main = Range("A1").Resize(-1, 1).Address', '-1'],
		];
		for (const [line, span] of cases) {
			const src = wrap(line);
			expectDiagnostic(src, analyzeModule(src), RANGE, { span, message: "'1004'" });
		}
		const quiet = wrap('Main = Range("B2").Offset(-1, -1).Address', 'Main = Range("A1").Resize(1, 1).Address', 'Main = Cells(1, 1).Value');
		expect(byCode(analyzeModule(quiet), RANGE)).toHaveLength(0);
	});

	it('flags a row or column past the bottom or right edge of the sheet (issue #182)', () => {
		const cases: Array<[string, string, string]> = [
			['Main = Cells(1048577, 1).Row', '1048577, 1', 'row 1048577'],
			['Main = ActiveSheet.Cells(1048577, 1).Row', '1048577, 1', 'row 1048577'],
			['Main = Cells(1, 16385).Row', '1, 16385', 'column 16385'],
			['Main = Rows(1048577).Row', '1048577', 'row 1048577'],
			['Main = Columns(16385).Column', '16385', 'column 16385'],
			['Main = Range("A1048576").Offset(1, 0).Row', '1, 0', 'row 1048577'],
			['Main = Range("A1").Offset(0, 16384).Column', '0, 16384', 'column 16385'],
			['Main = Range("A2").Resize(1048576).Row', '1048576', 'row 1048577'],
			['Main = Range("B1").Resize(1, 16384).Row', '1, 16384', 'column 16385'],
			['Main = Range("B2").Cells(1048576, 1).Row', '1048576, 1', 'row 1048577'],
			['Main = Range("A5").Rows(1048573).Row', '1048573', 'row 1048577'],
			['Main = Range("A1:A3").Rows(1048577).Row', '1048577', 'row 1048577'],
		];
		for (const [line, span, where] of cases) {
			const src = wrap(line);
			expectDiagnostic(src, analyzeModule(src), RANGE, { span, message: [where, "'1004'"] });
		}
		const quiet = wrap(
			'Main = Cells(1048576, 16384).Row',
			'Main = Columns(16384).Column',
			'Main = Range("A1").Resize(1048576).Rows.Count',
			'Main = Range("A1").Offset(0, 16383).Column',
			'Main = Range("A1").Offset(1048575).Row',
			'Main = Range("A1").Cells(1048576, 1).Row',
			'Main = Range("A5").Rows(1048572).Row',
		);
		expect(byCode(analyzeModule(quiet), RANGE)).toHaveLength(0);
	});

	it('reads Offset and Resize arguments as an offset and a size, not an index (issue #182)', () => {
		// Offset(0) and Offset(-1) from A2 run. Resize(0) raises, for its own reason.
		const quiet = wrap('Main = Range("A2").Offset(0).Row', 'Main = Range("A2").Offset(-1).Row');
		expect(byCode(analyzeModule(quiet), RANGE)).toHaveLength(0);
		const src = wrap('Main = Range("A1").Resize(0).Row');
		expectDiagnostic(src, analyzeModule(src), RANGE, { span: '0', message: 'Resize needs at least one row and one column' });
	});

	it('flags an address literal off the sheet', () => {
		const bad = ['"$A$0"', '"A$0"', '"$A0"', '"Sheet1!$A$0"', '"0:0"', '"$A$1048577"', '"$XFE$1"', '"$XFE:$XFE"', '"A0:$A$0"', '"$A$0:D4"'];
		for (const literal of bad) {
			const src = wrap(`Main = Range(${literal}).Value`);
			expectDiagnostic(src, analyzeModule(src), RANGE, { span: literal, message: ['1048576', "'1004'"] });
		}
		const twoArgs = wrap('Main = Range("$A$0", "B2").Value');
		expectDiagnostic(twoArgs, analyzeModule(twoArgs), RANGE, { span: '"$A$0"' });
		const quiet = wrap('Main = Range("A:A").Count', 'Main = Range("XFD1048576").Row', 'Main = Range("MyName").Value', 'Main = Application.Range("A1").Value');
		expect(byCode(analyzeModule(quiet), RANGE)).toHaveLength(0);
	});

	// Names.Add accepts each of these as a workbook name, and Range then
	// finds it (measured in Excel 16.0, 2026-10-01). A name cannot hold `$`
	// or start with a digit, so those addresses stay reported above.
	it('leaves an address a workbook name can spell alone', () => {
		for (const literal of ['"A0"', '"Sheet1!A0"', '"A1048577"', '"XFE1"', '"XFE:XFE"', '"A0:D4"', '"A1:XFE1"', '"A1048577:A1048577"', '"A0:$A$1"']) {
			const src = wrap(`Main = Range(${literal}).Value`);
			expect(byCode(analyzeModule(src), RANGE), literal).toHaveLength(0);
		}
		const twoArgs = wrap('Main = Range("A0", "D4").Value');
		expect(byCode(analyzeModule(twoArgs), RANGE)).toHaveLength(0);
	});

	it('leaves a name the module declares alone', () => {
		const src = 'Option Explicit\nFunction Cells(r As Long, c As Long) As Long\n    Cells = r\nEnd Function\nFunction Main() As Variant\n    Main = Cells(0, 1)\nEnd Function\n';
		expect(byCode(analyzeModule(src), RANGE)).toHaveLength(0);
	});
});

// Measured in Excel 16.0 (build 20326, 2026-10-01).
describe('host-argument-out-of-range - indexes relative to a range (issue #275)', () => {
	it('stays quiet where the index lands on the sheet', () => {
		for (const expr of [
			'Range("B2").Cells(0)', 'Range("B2").Cells(0, 0)', 'Range("B2").Cells(0, 1)', 'Range("C3").Cells(-1, -1)',
			'Range("B2").Item(0)', 'Range("B2").Item(0, 0)', 'Range("B2:C3").Rows(0)', 'Range("B2:C3").Columns(0)',
			'Range("B2:C3").Cells(0)', 'Range("B2:C3").Cells(-1)', 'Range("B2:D3").Cells(0)', 'Range("B2:D3").Cells(-2)',
			'Range("C3").Cells(-1)', 'Range("A2").Cells(0)', 'Range("B2:C3").Item(0)', 'Range("B1:C2").Cells(0)',
			'Range("B3:C4").Cells(-4)', 'Range("A2").Cells(0, 1)', 'Range("B2").Cells(1, 0)', 'Range("B2:C3").Cells(0, 0)',
			'Range("B3:C4").Rows(-1)', 'Range("B2").Columns(0)', 'Worksheets(1).Range("B2").Cells(0)', 'ActiveCell.Cells(0)',
		]) {
			expect(byCode(analyzeModule(wrap(`Main = ${expr}.Address`)), RANGE), expr).toEqual([]);
		}
		expect(byCode(analyzeModule(wrap('Dim r As Range', 'Set r = Range("B2")', 'Main = r.Cells(0, 1).Address')), RANGE)).toEqual([]);
	});

	it('reports where the index lands above row 1 or left of column A', () => {
		for (const [expr, where] of [
			['Range("A1").Cells(0)', 'row 0'],
			['Range("A1").Cells(0, 1)', 'row 0'],
			['Range("B2").Cells(-1)', 'row 0'],
			['Range("C3").Cells(-2)', 'row 0'],
			['Range("B1").Cells(0)', 'row 0'],
			['Range("A2:B3").Cells(0)', 'column 0'],
			['Range("B3:C4").Cells(-5)', 'row 0'],
			['Range("A1").Item(0, 1)', 'row 0'],
			['Range("A2").Cells(-1, 1)', 'row 0'],
			['Range("A1").Cells(1, 0)', 'column 0'],
			['Range("B2:C3").Cells(-1, 0)', 'row 0'],
			['Range("A1").Rows(0)', 'row 0'],
			['Range("A1:B2").Columns(0)', 'column 0'],
			['Range("B2:C3").Rows(-1)', 'row 0'],
			['Range("B2").Columns(-1)', 'column 0'],
			['Worksheets(1).Range("A1").Cells(0)', 'row 0'],
		]) {
			const src = wrap(`Main = ${expr}.Address`);
			expectDiagnostic(src, byCode(analyzeModule(src), RANGE), RANGE, { message: where });
		}
	});

	it('still reports the sheet\'s own Cells, Rows and Columns', () => {
		for (const expr of ['Cells(0, 1)', 'Rows(0)', 'Columns(0)', 'Worksheets(1).Cells(0, 1)', 'Worksheets(1).Rows(0)', 'Range("B2:C3").Areas(0)']) {
			expect(byCode(analyzeModule(wrap(`Main = ${expr}.Address`)), RANGE), expr).toHaveLength(1);
		}
	});
});

describe('host-argument-out-of-range - Word and PowerPoint (issue #122)', () => {
	it('flags Word collections at 0 and a Document.Range below 0', () => {
		const model = getWordObjectModel();
		const cases: Array<[string, string]> = [
			['Main = ActiveDocument.Paragraphs(0).Range.Text', "'5941'"],
			['Main = Documents(0).Name', "'5941'"],
			['Main = ActiveDocument.Tables(0).Rows.Count', "'5941'"],
			['Main = ActiveDocument.Paragraphs.Item(0).Range.Text', "'5941'"],
			['Main = ActiveDocument.Shapes(0).Name', "'-2147024809'"],
			['Main = ActiveDocument.Range(-1, 0).Text', "'4608'"],
			['Main = ActiveDocument.Range(1, 0).Text', "'4608'"],
		];
		for (const [line, error] of cases) {
			const src = wrap(line);
			const hits = byCode(analyzeModule(src, { hostModel: model }), RANGE);
			expect(hits, line).toHaveLength(1);
			expect(hits[0].message, line).toContain(error);
		}
		const quiet = wrap('Main = ActiveDocument.Range(0, 0).Text', 'Main = ActiveDocument.Paragraphs(1).Range.Text');
		expect(byCode(analyzeModule(quiet, { hostModel: model }), RANGE)).toHaveLength(0);
	});

	it('flags PowerPoint collections at 0 and Slides.Add at 0', () => {
		const model = getPowerPointObjectModel();
		const cases = [
			'Main = ActivePresentation.Slides(0).Name',
			'Main = Presentations(0).Name',
			'Main = ActivePresentation.Slides.Item(0).Name',
			'ActivePresentation.Slides.Add 0, ppLayoutBlank',
			'ActivePresentation.Slides.Add -1, ppLayoutBlank',
		];
		for (const line of cases) {
			const src = wrap(line);
			const hits = byCode(analyzeModule(src, { hostModel: model }), RANGE);
			expect(hits, line).toHaveLength(1);
			expect(hits[0].message, line).toContain("'-2147188160'");
		}
		const quiet = wrap('ActivePresentation.Slides.Add 1, ppLayoutBlank', 'Main = ActivePresentation.Slides(1).Name');
		expect(byCode(analyzeModule(quiet, { hostModel: model }), RANGE)).toHaveLength(0);
	});
});

describe('sheet-name-invalid (issue #122)', () => {
	it('flags a blank, over-long or forbidden-character sheet name', () => {
		const cases: Array<[string, string]> = [
			['Worksheets(1).Name = "a:b"', 'cannot contain'],
			['Worksheets(1).Name = "a[b"', 'cannot contain'],
			['Worksheets(1).Name = "a\\b"', 'cannot contain'],
			['Worksheets(1).Name = ""', 'blank'],
			['Worksheets(1).Name = "abcdefghijklmnopqrstuvwxyz123456"', '32'],
			['ActiveChart.Name = "a?b"', 'cannot contain'],
		];
		for (const [line, message] of cases) {
			const src = wrap(line);
			expectDiagnostic(src, analyzeModule(src), NAME, { message: [message, "'1004'"] });
		}
		const quiet = wrap('Worksheets(1).Name = "abcdefghijklmnopqrstuvwxyz12345"', 'Worksheets(1).Name = "Data 2026"', 'Names(1).Name = "a:b"');
		expect(byCode(analyzeModule(quiet), NAME)).toHaveLength(0);
	});

	// Measured in Excel 16.0 (build 20326, 2026-10-01).
	it('flags History, an apostrophe at either end, and a name String$ or Space$ spells out (issue #276)', () => {
		const cases: Array<[string, string, string]> = [
			['Worksheets(1).Name = "History"', 'History', 'History is a reserved name.'],
			['ActiveSheet.Name = "hIsToRy"', 'History', 'History is a reserved name.'],
			['ActiveChart.Name = "HISTORY"', 'History', 'History is a reserved name.'],
			['Worksheets(1).Name = "\'a"', 'apostrophe', 'invalid name'],
			['Worksheets(1).Name = "a\'"', 'apostrophe', 'invalid name'],
			['Worksheets(1).Name = "\' a"', 'apostrophe', 'invalid name'],
			['Worksheets(1).Name = String$(32, "a")', '32', 'invalid name'],
			['Worksheets(1).Name = String(32, "ab")', '32', 'invalid name'],
			['Worksheets(1).Name = Space$(32)', '32', 'invalid name'],
			['Worksheets(1).Name = String$(30, "a") & "bc"', '32', 'invalid name'],
		];
		for (const [line, message, error] of cases) {
			const src = wrap(line);
			expectDiagnostic(src, analyzeModule(src), NAME, { message: [message, error] });
		}
		const quiet = wrap(
			'Worksheets(1).Name = "History "',
			'Worksheets(1).Name = "History1"',
			'Worksheets(1).Name = "My History"',
			'Worksheets(1).Name = "a\'b"',
			'Worksheets(1).Name = " "',
			'Worksheets(1).Name = String$(31, "a")',
			'Worksheets(1).Name = Space(31)',
			'Worksheets(1).Name = "a-b c"',
		);
		expect(byCode(analyzeModule(quiet), NAME)).toHaveLength(0);
		// A String or Space the module declares is the module's own.
		const own = 'Option Explicit\nFunction Space(n As Long) As String\n    Space = "x"\nEnd Function\nFunction Main() As Variant\n    Worksheets(1).Name = Space(32)\nEnd Function\n';
		expect(byCode(analyzeModule(own), NAME)).toHaveLength(0);
	});
});

describe('host-argument-out-of-range - blank addresses and shared sheet members (issue #276)', () => {
	it('flags Range("") and Range(" ")', () => {
		for (const expr of ['Range("")', 'Range(" ")']) {
			const src = wrap(`Main = ${expr}.Address`);
			expectDiagnostic(src, byCode(analyzeModule(src), RANGE), RANGE, { message: ['blank', "'1004'"] });
		}
	});

	it('flags Shapes(0) on ActiveSheet and Sheets(1), which a Worksheet and a Chart both have', () => {
		for (const expr of ['ActiveSheet.Shapes(0)', 'Sheets(1).Shapes(0)', 'ActiveSheet.Shapes(-1)']) {
			const src = wrap(`Main = ${expr}.Name`);
			expectDiagnostic(src, byCode(analyzeModule(src), RANGE), RANGE, { message: "'-2147024809'" });
		}
	});

	it('does not judge an address a workbook name may spell', () => {
		for (const expr of ['Range("A1:ZZZZ1")', 'Range("ZZZZ1")', 'Range("AAAA1")']) {
			expect(byCode(analyzeModule(wrap(`Main = ${expr}.Address`)), RANGE), expr).toEqual([]);
		}
	});
});

describe('multi-cell-range-as-scalar (issue #122)', () => {
	it('flags a multi-cell literal assigned to a scalar variable, with or without .Value', () => {
		for (const type of ['String', 'Long', 'Integer', 'Double', 'Boolean', 'Date']) {
			const src = wrap(`Dim s As ${type}`, 's = Range("A1:B2")', 'Main = s');
			expectDiagnostic(src, analyzeModule(src), SCALAR, { span: 'Range("A1:B2")', message: [type, "'13'"] });
		}
		const value = wrap('Dim s As String', 's = Range("A1:B2").Value', 'Main = s');
		expectDiagnostic(value, analyzeModule(value), SCALAR, { span: 'Range("A1:B2").Value' });
		const qualified = wrap('Dim s As String', 's = ThisWorkbook.Worksheets(1).Range("A1:B2")', 'Main = s');
		expectDiagnostic(qualified, analyzeModule(qualified), SCALAR, { span: 'ThisWorkbook.Worksheets(1).Range("A1:B2")' });
	});

	it('flags a multi-cell literal beside a scalar operator', () => {
		const cases: Array<[string, string]> = [
			['If Range("A1:B2") = 5 Then Main = 1', '='],
			['If Range("A1:B2") < 5 Then Main = 1', '<'],
			['Main = Range("A1:B2") + 1', '+'],
			['Main = Range("A1:B2") & "x"', '&'],
		];
		for (const [line, operator] of cases) {
			const src = wrap(line);
			expectDiagnostics(src, analyzeModule(src), SCALAR, [{ span: 'Range("A1:B2")', message: `'${operator}'` }]);
		}
	});

	it('stays quiet for a single cell, a Variant target, a member and an index', () => {
		const src = wrap(
			'Dim s As String, v As Variant',
			's = Range("A1")',
			'v = Range("A1:B2")',
			'Main = Range("A1:B2").Address',
			'Main = Range("A1:B2")(1, 1)',
			'Main = Range("A1:B2").Value(1, 1)',
			'Main = s',
		);
		expect(byCode(analyzeModule(src), SCALAR)).toHaveLength(0);
	});
});

describe('a range in a one-line If branch (issue #140)', () => {
	// Measured in Excel 16.0 (build 20326, 2026-09-26): `If r Is Nothing Then
	// Set r = ws.Range("A1:P36")` runs; the range is the Set's value, not a
	// scalar operand of the condition.
	it('does not judge a Set after Then as a comparison', () => {
		const src = wrap('Dim ws As Worksheet, r As Range', 'Set ws = ActiveSheet', 'If r Is Nothing Then Set r = ws.Range("A1:P36")', 'Main = r.Cells.Count');
		expect(byCode(analyzeModule(src), SCALAR)).toHaveLength(0);
	});

	it('still judges a range compared in the condition itself', () => {
		const src = wrap('Dim ws As Worksheet', 'Set ws = ActiveSheet', 'If ws.Range("A1:B2") = 1 Then Main = 2');
		expectDiagnostic(src, analyzeModule(src), SCALAR, { message: 'Range("A1:B2")' });
	});
});
