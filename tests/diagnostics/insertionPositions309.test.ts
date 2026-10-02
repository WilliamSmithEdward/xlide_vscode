// Diagnostics tests: positions past a collection's end or at 0, written
// against its own Count (issue #309). Each sample was run through
// pyVBAharness on 2026-10-02 in Excel, Word and PowerPoint 16.0.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { getPowerPointObjectModel } from '../../src/analyzer/host/powerpointObjectModel';
import { getWordObjectModel } from '../../src/analyzer/host/wordObjectModel';

const LO = 'Dim lo As ListObject\n    Set lo = ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1:B3"), , xlYes)\n    ';
const DOC = 'Dim d As Document\n    Set d = Documents.Add\n    ';
const LAYOUT = 'ActivePresentation.SlideMaster.CustomLayouts(1)';

function errors(body: string, host?: 'Word' | 'PowerPoint'): string[] {
	const hostModel = host === 'Word' ? getWordObjectModel() : host === 'PowerPoint' ? getPowerPointObjectModel() : undefined;
	const source = `Option Explicit\nFunction Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(source, hostModel ? { hostModel } : {})
		.filter((diag) => diag.code === 'host-argument-out-of-range')
		.map((diag) => /Run-time error '(-?\d+)'/.exec(diag.message)?.[1] ?? '?');
}

describe('a position the collection does not reach (issue #309)', () => {
	it('is refused past Count + 1 or at 0 on an insertion', () => {
		expect(errors('ActivePresentation.Slides.Add ActivePresentation.Slides.Count + 2, ppLayoutBlank', 'PowerPoint')).toEqual(['-2147188160']);
		expect(errors('Dim p As Presentation\n    Set p = ActivePresentation\n    p.Slides.Add p.Slides.Count + 2, ppLayoutBlank', 'PowerPoint')).toEqual(['-2147188160']);
		expect(errors(`ActivePresentation.Slides.AddSlide 0, ${LAYOUT}`, 'PowerPoint')).toEqual(['-2147188160']);
		expect(errors(`ActivePresentation.Slides.AddSlide ActivePresentation.Slides.Count + 2, ${LAYOUT}`, 'PowerPoint')).toEqual(['-2147188160']);
		expect(errors('ActivePresentation.Slides.Add 0, ppLayoutBlank', 'PowerPoint')).toEqual(['-2147188160']);
		expect(errors(`${LO}lo.ListRows.Add 0`)).toEqual(['9']);
		expect(errors(`${LO}lo.ListColumns.Add 0`)).toEqual(['9']);
		expect(errors(`${LO}lo.ListRows.Add lo.ListRows.Count + 2`)).toEqual(['9']);
		expect(errors(`${LO}lo.ListColumns.Add lo.ListColumns.Count + 2`)).toEqual(['9']);
	});

	it('is refused past Count on a read', () => {
		expect(errors('Main = Worksheets(Worksheets.Count + 1).Name')).toEqual(['9']);
		expect(errors(`${DOC}Main = d.Paragraphs(d.Paragraphs.Count + 1).Range.Text`, 'Word')).toEqual(['5941']);
	});

	it('stays quiet at Count + 1 on an insertion, at Count on a read, and on another collection\'s Count', () => {
		expect(errors('ActivePresentation.Slides.Add ActivePresentation.Slides.Count + 1, ppLayoutBlank', 'PowerPoint')).toEqual([]);
		expect(errors(`ActivePresentation.Slides.AddSlide 1, ${LAYOUT}`, 'PowerPoint')).toEqual([]);
		expect(errors(`${LO}lo.ListRows.Add lo.ListRows.Count + 1\n    lo.ListColumns.Add 1`)).toEqual([]);
		expect(errors('Main = Worksheets(Worksheets.Count).Name')).toEqual([]);
		expect(errors(`${DOC}Main = d.Paragraphs(d.Paragraphs.Count).Range.Text`, 'Word')).toEqual([]);
		expect(errors('Main = Worksheets(Workbooks.Count + 1).Name')).toEqual([]);
		expect(errors('Main = Range("A1:A3").Rows(Range("A1:A3").Rows.Count + 1).Address')).toEqual([]);
	});
});
