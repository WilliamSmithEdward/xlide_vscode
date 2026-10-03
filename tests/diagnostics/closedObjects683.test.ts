// Diagnostics tests: what a closed workbook, document or presentation, or a
// deleted slide, leaves behind (issue #683). Measured on 2026-10-03 in Excel,
// Word and PowerPoint 16.0 through pyVBAharness.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { getPowerPointObjectModel } from '../../src/analyzer/host/powerpointObjectModel';
import { getWordObjectModel } from '../../src/analyzer/host/wordObjectModel';

function found(body: string, host?: 'Word' | 'PowerPoint'): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    ${body}\nEnd Function\n`;
	const hostModel = host === 'Word' ? getWordObjectModel() : host === 'PowerPoint' ? getPowerPointObjectModel() : undefined;
	return analyzeModule(src, hostModel ? { hostModel } : {})
		.filter((diag) => diag.code === 'object-used-after-delete')
		.map((diag) => diag.message);
}

describe('what is taken from an Excel workbook before it closes (issue #683)', () => {
	it.each([
		['a Range of one of its sheets', 'Dim wb As Workbook, r As Range\n    Set wb = Workbooks.Add\n    Set r = wb.Sheets(1).Range("A1")\n    wb.Close False\n    Main = r.Address', "'r' is a range of 'wb', which was closed on line 6, so its Address is gone. This will raise Run-time error '424': Object required."],
		['a Range through Cells', 'Dim wb As Workbook, r As Range\n    Set wb = Workbooks.Add\n    Set r = wb.Worksheets(1).Cells(1, 1)\n    wb.Close False\n    Main = r.Value', "Run-time error '424'"],
		['a Range of ActiveSheet', 'Dim wb As Workbook, r As Range\n    Set wb = Workbooks.Add\n    Set r = wb.ActiveSheet.Range("A1")\n    wb.Close False\n    Main = r.Address', "Run-time error '424'"],
		['a sheet', 'Dim wb As Workbook, ws As Worksheet\n    Set wb = Workbooks.Add\n    Set ws = wb.Sheets(1)\n    wb.Close False\n    Main = ws.Name', "'ws' is a sheet of 'wb', which was closed on line 6, so its Name is gone. This will raise Run-time error '-2147221080': Method 'Name' of object '_Worksheet' failed."],
		['a Range of a sheet of it', 'Dim wb As Workbook, ws As Worksheet, r As Range\n    Set wb = Workbooks.Add\n    Set ws = wb.Worksheets(1)\n    Set r = ws.Range("A1")\n    wb.Close False\n    Main = r.Address', "'r' is a range of 'wb'"],
	])('reports %s', (_label, body, message) => {
		const messages = found(body);
		expect(messages, body).toHaveLength(1);
		expect(messages[0], body).toContain(message);
	});

	it('stays quiet on Is Nothing and a new workbook', () => {
		expect(found('Dim wb As Workbook\n    Set wb = Workbooks.Add\n    wb.Close False\n    Main = (wb Is Nothing)')).toEqual([]);
		expect(found('Dim wb As Workbook\n    Set wb = Workbooks.Add\n    wb.Close False\n    Set wb = Workbooks.Add\n    Main = wb.Name')).toEqual([]);
		expect(found('Dim wb As Workbook, r As Range\n    Set wb = Workbooks.Add\n    Set r = wb.Sheets(1).Range("A1")\n    Set r = ActiveSheet.Range("A1")\n    wb.Close False\n    Main = r.Address')).toEqual([]);
	});
});

describe('a closed Word document (issue #683)', () => {
	it.each([
		['a member', 'Dim d As Document\n    Set d = Documents.Add\n    d.Close SaveChanges:=False\n    Main = d.Name', "'d' was closed on line 5, so its Name is gone. This will raise Run-time error '5825': Object has been deleted."],
		['its Content', 'Dim d As Document\n    Set d = Documents.Add\n    d.Close False\n    Main = d.Content.Text', "so its Content is gone"],
		['a Range of its Content', 'Dim d As Document, r As Range\n    Set d = Documents.Add\n    Set r = d.Content\n    d.Close False\n    Main = r.Text', "'r' is a range of 'd', which was closed on line 6, so its Text is gone. This will raise Run-time error '5825': Object has been deleted."],
		['a Range of it', 'Dim d As Document, r As Range\n    Set d = Documents.Add\n    Set r = d.Range\n    d.Close False\n    Main = r.Text', "Run-time error '5825'"],
		['a paragraph\'s Range', 'Dim d As Document, r As Range\n    Set d = Documents.Add\n    Set r = d.Paragraphs(1).Range\n    d.Close False\n    Main = r.Text', "Run-time error '5825'"],
	])('reports %s', (_label, body, message) => {
		const messages = found(body, 'Word');
		expect(messages, body).toHaveLength(1);
		expect(messages[0], body).toContain(message);
	});

	it('stays quiet on Is Nothing and a new document', () => {
		expect(found('Dim d As Document\n    Set d = Documents.Add\n    d.Close 0\n    Main = (d Is Nothing)', 'Word')).toEqual([]);
		expect(found('Dim d As Document\n    Set d = Documents.Add\n    d.Close False\n    Set d = Documents.Add\n    Main = d.Name', 'Word')).toEqual([]);
	});
});

describe('a closed PowerPoint presentation and a deleted slide (issue #683)', () => {
	it.each([
		['a member', 'Dim p As Presentation\n    Set p = Presentations.Add(msoFalse)\n    p.Close\n    Main = p.Name', "'p' was closed on line 5, so its Name is gone. This will raise Run-time error '-2147188720': Presentation (unknown member) : Object does not exist."],
		['its Slides', 'Dim p As Presentation\n    Set p = Presentations.Add(msoFalse)\n    p.Close\n    Main = p.Slides.Count', "so its Slides is gone"],
		['a deleted slide', 'Dim p As Presentation, s As Slide\n    Set p = Presentations.Add(msoFalse)\n    Set s = p.Slides.Add(1, ppLayoutBlank)\n    s.Delete\n    Main = s.Name', "'s' was deleted on line 6, so its Name is gone. This will raise Run-time error '-2147188720': Slide (unknown member) : Object does not exist."],
		['a slide it added', 'Dim p As Presentation, s As Slide\n    Set p = Presentations.Add(msoFalse)\n    Set s = p.Slides.Add(1, ppLayoutBlank)\n    p.Close\n    Main = s.SlideIndex', "'s' is a slide of 'p', which was closed on line 6"],
		['a slide of it', 'Dim p As Presentation, s As Slide\n    Set p = Presentations.Add(msoFalse)\n    p.Slides.Add 1, ppLayoutBlank\n    Set s = p.Slides(1)\n    p.Close\n    Main = s.Name', "'s' is a slide of 'p'"],
	])('reports %s', (_label, body, message) => {
		const messages = found(body, 'PowerPoint');
		expect(messages, body).toHaveLength(1);
		expect(messages[0], body).toContain(message);
	});

	it('stays quiet on Is Nothing and a new slide', () => {
		expect(found('Dim p As Presentation\n    Set p = Presentations.Add(msoFalse)\n    p.Close\n    Main = (p Is Nothing)', 'PowerPoint')).toEqual([]);
		expect(found('Dim p As Presentation, s As Slide\n    Set p = Presentations.Add(msoFalse)\n    Set s = p.Slides.Add(1, ppLayoutBlank)\n    s.Delete\n    Set s = p.Slides.Add(1, ppLayoutBlank)\n    Main = s.SlideIndex\n    p.Close', 'PowerPoint')).toEqual([]);
	});
});
