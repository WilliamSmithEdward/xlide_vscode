// Assignments to read-only host properties (issue #198). Every case was
// measured through pyVBAharness: Excel, Word and PowerPoint 16.0 compile each
// statement below, and Access the two Access ones.

import { describe, it, expect } from 'vitest';
import { hostTokenForFileName } from '../../src/analyzer/host/hostRegistry';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';

const CODE = 'readonly-member-assignment';
const READ_ONLY = "Can't assign to read-only property";
const WRONG_NUMBER = 'Wrong number of arguments or invalid property assignment';
const CONSTANT = 'Assignment to constant not permitted';
const INVALID_USE = 'Invalid use of property';

function readOnlyMessages(file: string, decls: string, stmt: string): string[] {
	const source = `Option Explicit\nSub S()\n${[decls, stmt].filter(Boolean).map((line) => `    ${line}`).join('\n')}\nEnd Sub\n`;
	return analyzeVbaModuleSource({
		source,
		moduleName: 'Module1',
		host: hostTokenForFileName(file),
		referencedHosts: [],
	}).diagnostics.filter((d) => d.code === CODE).map((d) => d.message);
}

describe('a read-only host property is not assigned (issue #198)', () => {
	const REFUSED: ReadonlyArray<readonly [string, string, string, string]> = [
		// The dual interfaces: Excel's Application, Workbook, Worksheet, Chart,
		// Workbooks, Sheets and drawing formats, and all of Word, PowerPoint,
		// Access and Office.
		['Book.xlsm', '', 'ActiveWorkbook.Name = "x"', READ_ONLY],
		['Book.xlsm', 'Dim wb As Workbook', 'wb.Name = "x"', READ_ONLY],
		['Book.xlsm', '', 'Application.Version = "1"', READ_ONLY],
		['Book.xlsm', 'Dim ws As Worksheet', 'ws.Index = 2', READ_ONLY],
		['Book.xlsm', 'Dim wb As Workbook', 'wb.FullName = "x"', READ_ONLY],
		['Book.xlsm', 'Dim wbs As Workbooks', 'wbs.Count = 2', READ_ONLY],
		['Book.xlsm', 'Dim s As Sheets', 's.Count = 2', READ_ONLY],
		['Book.xlsm', 'Dim c As Chart', 'c.Index = 2', READ_ONLY],
		['Book.xlsm', 'Dim f As FillFormat', 'f.Type = 2', READ_ONLY],
		['Book.xlsm', '', 'ThisWorkbook.Name = "x"', READ_ONLY],
		['Book.xlsm', 'Dim wb As Workbook', 'Let wb.Name = "x"', READ_ONLY],
		['Book.xlsm', 'Dim wb As Workbook', 'wb.Name() = "x"', READ_ONLY],
		['Book.xlsm', 'Dim wb As Workbook', 'With wb\n        .Name = "x"\n    End With', READ_ONLY],
		['Doc.docm', 'Dim d As Document', 'd.Name = "x"', READ_ONLY],
		['Doc.docm', 'Dim p As Paragraphs', 'p.Count = 2', READ_ONLY],
		['Doc.docm', '', 'Documents(1).Name = "x"', READ_ONLY],
		['Doc.docm', 'Dim r As Range', 'r.XML(False) = "x"', READ_ONLY],
		['Deck.pptm', 'Dim s As Slide', 's.SlideIndex = 2', READ_ONLY],
		['Deck.pptm', '', 'ActivePresentation.Slides(1).SlideIndex = 2', READ_ONLY],
		['Db.accdb', 'Dim c As Controls', 'c.Count = 1', READ_ONLY],
		['Db.accdb', 'Dim cp As CurrentProject', 'cp.Name = "x"', READ_ONLY],
		// Excel's dispatch-only interfaces.
		['Book.xlsm', 'Dim x As Range', 'x.Count = 5', WRONG_NUMBER],
		['Book.xlsm', '', 'Range("A1").Count = 5', WRONG_NUMBER],
		['Book.xlsm', 'Dim ws As Worksheet', 'ws.Range("A1").Row = 1', WRONG_NUMBER],
		['Book.xlsm', 'Dim s As Worksheets', 's.Count = 2', WRONG_NUMBER],
		['Book.xlsm', 'Dim sh As Shape', 'sh.ID = 2', WRONG_NUMBER],
		['Book.xlsm', 'Dim co As ChartObject', 'co.Index = 2', WRONG_NUMBER],
		['Book.xlsm', 'Dim r As Range', 'r.Address = "B2"', CONSTANT],
		['Book.xlsm', 'Dim r As Range', 'r.Address(False, False) = "B2"', CONSTANT],
		['Book.xlsm', 'Dim r As Range', 'r.AddressLocal = "B2"', CONSTANT],
		// Set onto a read-only property.
		['Book.xlsm', 'Dim ws As Worksheet, r As Range', 'Set ws.UsedRange = r', INVALID_USE],
		['Book.xlsm', 'Dim r As Range, a As Areas', 'Set r.Areas = a', INVALID_USE],
		['Book.xlsm', 'Dim wb As Workbook, o As Object', 'Set wb.Name = o', INVALID_USE],
		['Book.xlsm', 'Dim r As Range, o As Object', 'Set r.Count = o', WRONG_NUMBER],
		['Book.xlsm', 'Dim r As Range, o As Object', 'Set r.Address = o', CONSTANT],
		['Doc.docm', 'Dim d As Document, r As Range', 'Set d.Content = r', INVALID_USE],
		['Deck.pptm', 'Dim s As Slide, sh As Shapes', 'Set s.Shapes = sh', INVALID_USE],
	];
	it.each(REFUSED)('%s: %s ... %s is "%s"', (file, decls, stmt, error) => {
		expect(readOnlyMessages(file, decls, stmt)).toEqual([
			expect.stringContaining(`This is a VBE compile error: ${error}.`),
		]);
	});

	const COMPILES: ReadonlyArray<readonly [string, string, string]> = [
		// A Variant or Object property takes the value when the code runs.
		['Book.xlsm', 'Dim r As Range', 'r.Text = "a"'],
		['Book.xlsm', 'Dim r As Range', 'r.CountLarge = 2'],
		['Book.xlsm', 'Dim r As Range', 'r.HasArray = True'],
		['Book.xlsm', 'Dim ws As Worksheet', 'ws.Parent = 1'],
		['Book.xlsm', 'Dim r As Range, o As Object', 'Set r.Text = o'],
		['Book.xlsm', 'Dim ws As Worksheet, o As Object', 'Set ws.Parent = o'],
		// A Let to an object property goes to its default member.
		['Book.xlsm', 'Dim ws As Worksheet', 'ws.UsedRange = 1'],
		['Book.xlsm', 'Dim ws As Worksheet', 'ws.Rows = 1'],
		['Book.xlsm', 'Dim r As Range', 'r.Cells(1, 1) = 1'],
		// A late-bound receiver binds when the code runs.
		['Book.xlsm', '', 'ActiveSheet.Index = 2'],
		['Book.xlsm', '', 'Worksheets(1).Name = "x"'],
		['Book.xlsm', '', 'Sheets(1).Index = 2'],
		['Book.xlsm', 'Dim wb As Workbook', 'wb.Worksheets(1).Name = "x"'],
		['Book.xlsm', 'Dim o As Object', 'o.Name = "x"'],
		['Book.xlsm', 'Dim v As Variant', 'v.Name = "x"'],
		// Read-write, whatever the model's documentation says of Type.
		['Book.xlsm', 'Dim p As DocumentProperty', 'p.Type = 1'],
		// A comparison in a single-line If is no assignment.
		['Doc.docm', 'Dim p As Paragraph', 'With p.TabStops\n        If .Count = 0 Then .Add 18\n    End With'],
		['Book.xlsm', 'Dim r As Range', 'If r.Count = 0 Then Exit Sub'],
	];
	it.each(COMPILES)('%s: %s ... %s compiles', (file, decls, stmt) => {
		expect(readOnlyMessages(file, decls, stmt)).toEqual([]);
	});

	it('leaves a Set to a scalar Word property alone, whose error turns on parameters the model lacks', () => {
		// The VBE says "Type mismatch" for Set r.XML = o, where XML takes a
		// parameter, and "Invalid use of property" where the property takes none.
		expect(readOnlyMessages('Doc.docm', 'Dim r As Range, o As Object', 'Set r.XML = o')).toEqual([]);
	});

	it('leaves a method alone, which the VBE refuses for its arguments instead', () => {
		// "Argument not optional" here, and "Wrong number of arguments or
		// invalid property assignment" for Set los.Item(1) = lo.
		expect(readOnlyMessages('Book.xlsm', 'Dim los As ListObjects, lo As ListObject', 'Set los.Item = lo')).toEqual([]);
	});
});
