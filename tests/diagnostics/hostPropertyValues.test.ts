// Host property values and counts the host refuses (issue #204). Each case
// was measured through pyVBAharness in Excel, Word and PowerPoint 16.0: the
// reported ones raise the error named every time they run, the quiet ones run.

import { describe, it, expect } from 'vitest';
import { hostTokenForFileName } from '../../src/analyzer/host/hostRegistry';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';

const PPT_SETUP = 'Dim sld As Slide, shp As Shape\n    Set sld = ActivePresentation.Slides.Add(1, 12)\n    Set shp = sld.Shapes.AddTextbox(1, 10, 10, 100, 50)';

function hostErrors(file: string, body: string): string[] {
	const source = `Option Explicit\nSub S()\n    ${body}\nEnd Sub\n`;
	return analyzeVbaModuleSource({ source, moduleName: 'Module1', host: hostTokenForFileName(file), referencedHosts: [] })
		.diagnostics
		.filter((d) => d.code === 'host-property-value-out-of-range' || d.code === 'host-argument-out-of-range')
		.map((d) => `${d.code}: ${d.message}`);
}

describe('host property values out of range (issue #204)', () => {
	const REFUSED: ReadonlyArray<readonly [string, string, string]> = [
		['Book.xlsm', 'Range("A1").Font.Size = 500', "Font.Size takes 1 to 409.5; 500 is outside that. This will raise Run-time error '1004': Unable to set the Size property of the Font class."],
		['Book.xlsm', 'Range("A1").Font.Size = 0', "'1004'"],
		['Book.xlsm', 'Range("A1").Font.Size = 0.5', "'1004'"],
		['Book.xlsm', 'Range("A1").Font.Size = 409.6', "'1004'"],
		['Book.xlsm', 'Range("A1").Interior.ColorIndex = 57', "Interior.ColorIndex takes 1 to 56, or an xlColorIndex constant; 57 is outside that. This will raise Run-time error '9': Subscript out of range."],
		['Book.xlsm', 'ActiveSheet.Tab.ColorIndex = 0', "'9'"],
		['Book.xlsm', 'ActiveSheet.Tab.ColorIndex = -1', "'9'"],
		['Book.xlsm', 'ActiveSheet.Tab.ColorIndex = 99', "'9'"],
		['Book.xlsm', 'Rows(1).RowHeight = 500', "Unable to set the RowHeight property of the Range class"],
		['Book.xlsm', 'Rows(1).RowHeight = 409.75', "RowHeight"],
		['Book.xlsm', 'Rows(1).RowHeight = -0.5', "RowHeight"],
		['Book.xlsm', 'Columns(1).ColumnWidth = 255.5', "Unable to set the ColumnWidth property"],
		['Book.xlsm', 'Columns(1).ColumnWidth = -1', "ColumnWidth"],
		['Book.xlsm', 'ActiveWindow.Zoom = 9', "Unable to set the Zoom property of the Window class"],
		['Book.xlsm', 'ActiveWindow.Zoom = 0', "Zoom"],
		['Book.xlsm', 'ActiveWindow.Zoom = -5', "Zoom"],
		['Book.xlsm', 'ActiveWindow.Zoom = 401', "Zoom"],
		['Book.xlsm', 'Range("A1").Orientation = 91', "Unable to set the Orientation property"],
		['Book.xlsm', 'Range("A1").Orientation = -4000', "Orientation"],
		['Book.xlsm', 'Range("A1").IndentLevel = 251', "Unable to set the IndentLevel property"],
		['Book.xlsm', 'Range("A1").IndentLevel = -16', "IndentLevel"],
		['Book.xlsm', 'Dim f As Font\n    Set f = Range("A1").Font\n    f.Size = 500', "'1004'"],
		['Book.xlsm', 'With Range("A1").Font\n        .Size = 500\n    End With', "'1004'"],
		['Doc.docm', 'ActiveDocument.Content.Font.Size = 2000', "Font.Size takes 1 to 1638; 2000 is outside that. This will raise Run-time error '5843'"],
		['Doc.docm', 'ActiveDocument.Content.Font.Size = 0.5', "'5843'"],
		['Doc.docm', 'ActiveDocument.Content.Font.Size = 1638.5', "'5843'"],
		['Doc.docm', 'ActiveWindow.View.Zoom.Percentage = 5', "Zoom.Percentage takes 10 to 500"],
		['Doc.docm', 'ActiveWindow.View.Zoom.Percentage = 501', "'5843'"],
		['Doc.docm', 'ActiveDocument.Paragraphs(1).LeftIndent = 100000', "'5149': The measurement must be between -1584 pt and 1584 pt."],
		['Doc.docm', 'ActiveDocument.Paragraphs(1).LeftIndent = -1585', "'5149'"],
		['Deck.pptm', `${PPT_SETUP}\n    shp.TextFrame.TextRange.Font.Size = 5000`, "Font.Size takes 1 to 4000; 5000 is outside that. This will raise Run-time error '-2147024809': The specified value is out of range."],
		['Deck.pptm', `${PPT_SETUP}\n    shp.TextFrame.TextRange.Font.Size = 0`, "'-2147024809'"],
		['Deck.pptm', `${PPT_SETUP}\n    shp.TextFrame.TextRange.Font.Size = 4000.25`, "'-2147024809'"],
		// Counts a method refuses.
		['Book.xlsm', 'Worksheets.Add Count:=0', "host-argument-out-of-range: Worksheets.Add takes Count 1 or more; 0 is outside that. This will raise Run-time error '1004': Method 'Add' of object 'Sheets' failed."],
		['Book.xlsm', 'Worksheets.Add , , -1', "Count 1 or more; -1"],
		['Doc.docm', 'ActiveDocument.Tables.Add ActiveDocument.Range(0, 0), 0, 2', "Tables.Add takes NumRows 1 to 32767; 0 is outside that. This will raise Run-time error '5148': The number must be between 1 and 32767."],
		['Doc.docm', 'ActiveDocument.Tables.Add ActiveDocument.Range(0, 0), 32768, 1', "NumRows 1 to 32767; 32768"],
		['Doc.docm', 'ActiveDocument.Tables.Add ActiveDocument.Range(0, 0), 1, 64', "NumColumns 1 to 63; 64"],
		['Doc.docm', 'ActiveDocument.Tables.Add Range:=ActiveDocument.Range(0, 0), NumRows:=1, NumColumns:=0', "NumColumns 1 to 63; 0"],
	];
	it.each(REFUSED)('%s: %s', (file, body, text) => {
		const hits = hostErrors(file, body);
		expect(hits).toHaveLength(1);
		expect(hits[0]).toContain(text);
	});

	const RUNS: ReadonlyArray<readonly [string, string]> = [
		['Book.xlsm', 'Range("A1").Font.Size = 1'],
		['Book.xlsm', 'Range("A1").Font.Size = 409.5'],
		['Book.xlsm', 'Range("A1").Interior.ColorIndex = 56'],
		['Book.xlsm', 'Range("A1").Interior.ColorIndex = 0'],
		['Book.xlsm', 'Range("A1").Interior.ColorIndex = -1'],
		['Book.xlsm', 'Range("A1").Interior.ColorIndex = xlColorIndexNone'],
		['Book.xlsm', 'ActiveSheet.Tab.ColorIndex = 56'],
		['Book.xlsm', 'Rows(1).RowHeight = 0'],
		['Book.xlsm', 'Rows(1).RowHeight = 409.5'],
		['Book.xlsm', 'Columns(1).ColumnWidth = 255'],
		['Book.xlsm', 'ActiveWindow.Zoom = 10'],
		['Book.xlsm', 'ActiveWindow.Zoom = 400'],
		['Book.xlsm', 'ActiveWindow.Zoom = True'],
		['Book.xlsm', 'ActiveWindow.Zoom = False'],
		['Book.xlsm', 'ActiveWindow.Zoom = -1'],
		['Book.xlsm', 'Range("A1").Orientation = -90'],
		['Book.xlsm', 'Range("A1").Orientation = -4166'],
		['Book.xlsm', 'Range("A1").Orientation = xlUpward'],
		['Book.xlsm', 'Range("A1").IndentLevel = 250'],
		['Book.xlsm', 'Range("A1").IndentLevel = -1'],
		['Book.xlsm', 'Worksheets.Add Count:=1'],
		['Doc.docm', 'ActiveDocument.Content.Font.Size = 1638.25'],
		['Doc.docm', 'ActiveWindow.View.Zoom.Percentage = 500'],
		['Doc.docm', 'ActiveDocument.Paragraphs(1).LeftIndent = 1584'],
		['Doc.docm', 'ActiveDocument.Tables.Add ActiveDocument.Range(0, 0), 1, 63'],
		['Deck.pptm', `${PPT_SETUP}\n    shp.TextFrame.TextRange.Font.Size = 4000`],
	];
	it.each(RUNS)('%s: %s runs', (file, body) => {
		expect(hostErrors(file, body)).toEqual([]);
	});
});
