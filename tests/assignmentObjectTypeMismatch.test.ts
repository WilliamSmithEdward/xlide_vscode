import { describe, it, expect } from 'vitest';
import { analyzeVbaModuleSource } from '../src/vbaModuleAnalysis';

const MISMATCH = 'assignment-object-type-mismatch';

function codes(src: string): string[] {
	return analyzeVbaModuleSource({ source: src, moduleName: 'Module1' }).diagnostics.map((d) => d.code);
}

describe('assignment-object-type-mismatch: indexed collection element', () => {
	it('does not flag ThisWorkbook.Sheets("x") assigned to a Worksheet', () => {
		// Sheets(i) is a late-bound single sheet (Worksheet OR Chart); VBA allows
		// `Set ws = Sheets("x")`. Regression for the returnsAnyOf Item resolver.
		const src = 'Sub S()\n    Dim ws As Worksheet\n    Set ws = ThisWorkbook.Sheets("mySheet")\nEnd Sub\n';
		expect(codes(src)).not.toContain(MISMATCH);
	});

	it('does not flag ThisWorkbook.Worksheets("x") assigned to a Worksheet', () => {
		const src = 'Sub S()\n    Dim ws As Worksheet\n    Set ws = ThisWorkbook.Worksheets("mySheet")\nEnd Sub\n';
		expect(codes(src)).not.toContain(MISMATCH);
	});

	it('still flags a provable mismatch (Range assigned to a Worksheet)', () => {
		const src = 'Sub S()\n    Dim ws As Worksheet\n    Dim r As Range\n    Set ws = r\nEnd Sub\n';
		expect(codes(src)).toContain(MISMATCH);
	});
});

describe('indexed collection accessors resolve to the element type, not the collection', () => {
	// The whole `Collection([Index])` family - method-modelled (ChartObjects,
	// OLEObjects, Buttons, Pictures) and property-modelled-with-signature
	// (PivotFields, PivotItems) accessors alike - must resolve to a single element.
	const CASES: ReadonlyArray<readonly [string, string]> = [
		['Dim ws As Worksheet\n    Dim c As ChartObject', 'Set c = ws.ChartObjects(1)'],
		['Dim ws As Worksheet\n    Dim o As OLEObject', 'Set o = ws.OLEObjects(1)'],
		['Dim ws As Worksheet\n    Dim b As Button', 'Set b = ws.Buttons(1)'],
		['Dim ws As Worksheet\n    Dim p As Picture', 'Set p = ws.Pictures(1)'],
		['Dim ws As Worksheet\n    Dim pt As PivotTable', 'Set pt = ws.PivotTables(1)'],
		['Dim pt As PivotTable\n    Dim pf As PivotField', 'Set pf = pt.PivotFields(1)'],
		['Dim pf As PivotField\n    Dim pi As PivotItem', 'Set pi = pf.PivotItems(1)'],
		['Dim ch As Chart\n    Dim sr As Series', 'Set sr = ch.SeriesCollection(1)'],
		['Dim ch As Chart\n    Dim cg As ChartGroup', 'Set cg = ch.ChartGroups(1)'],
		['Dim sr As Series\n    Dim p As Point', 'Set p = sr.Points(1)'],
		['Dim sr As Series\n    Dim t As Trendline', 'Set t = sr.Trendlines(1)'],
	];
	it.each(CASES)('does not flag mismatch: %s ... %s', (decls, stmt) => {
		const src = `Sub S()\n    ${decls}\n    ${stmt}\nEnd Sub\n`;
		expect(codes(src)).not.toContain(MISMATCH);
	});

	it('chains through a method-kind indexed accessor (ws.ChartObjects(1).Chart)', () => {
		const src = 'Sub S()\n    Dim ws As Worksheet\n    Dim cc As Chart\n    Set cc = ws.ChartObjects(1).Chart\nEnd Sub\n';
		const cs = codes(src);
		expect(cs).not.toContain(MISMATCH);
		expect(cs).not.toContain('member-not-found');
	});

	it('keeps concrete-typed (non-collection) calls unchanged', () => {
		const range = 'Sub S()\n    Dim ws As Worksheet\n    Dim r As Range\n    Set r = ws.Range("A1")\nEnd Sub\n';
		const inter = 'Sub S()\n    Dim ws As Worksheet\n    Dim r As Range\n    Set r = Application.Intersect(ws.Range("A1"), ws.Range("B2"))\nEnd Sub\n';
		expect(codes(range)).not.toContain(MISMATCH);
		expect(codes(inter)).not.toContain(MISMATCH);
	});
});
