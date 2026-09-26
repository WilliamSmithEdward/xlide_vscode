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

	it('flags an address literal off the sheet', () => {
		const bad = ['"A0"', '"$A$0"', '"Sheet1!A0"', '"0:0"', '"A1048577"', '"XFE1"'];
		for (const literal of bad) {
			const src = wrap(`Main = Range(${literal}).Value`);
			expectDiagnostic(src, analyzeModule(src), RANGE, { span: literal, message: ['1048576', "'1004'"] });
		}
		const twoArgs = wrap('Main = Range("A0", "B2").Value');
		expectDiagnostic(twoArgs, analyzeModule(twoArgs), RANGE, { span: '"A0"' });
		const quiet = wrap('Main = Range("A:A").Count', 'Main = Range("XFD1048576").Row', 'Main = Range("MyName").Value', 'Main = Application.Range("A1").Value');
		expect(byCode(analyzeModule(quiet), RANGE)).toHaveLength(0);
	});

	it('leaves a name the module declares alone', () => {
		const src = 'Option Explicit\nFunction Cells(r As Long, c As Long) As Long\n    Cells = r\nEnd Function\nFunction Main() As Variant\n    Main = Cells(0, 1)\nEnd Function\n';
		expect(byCode(analyzeModule(src), RANGE)).toHaveLength(0);
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
