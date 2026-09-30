import { describe, it, expect } from 'vitest';
import { analyzeVbaModuleSource } from '../src/vbaModuleAnalysis';
import { hostTokenForFileName } from '../src/analyzer/host/hostRegistry';

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

describe('collection-of-collections accessors are not over-resolved', () => {
	// SparklineGroup is itself a collection (its Item is a Sparkline). Item/_Default/
	// Add already return the resolved element/result, so re-indexing them would
	// over-resolve one level too far (SparklineGroup -> Sparkline).
	it('does not flag SparklineGroups.Add(...) assigned to a SparklineGroup', () => {
		const src = 'Sub S()\n    Dim rng As Range\n    Dim sg As SparklineGroup\n    Set sg = rng.SparklineGroups.Add(1, "A1:A5")\nEnd Sub\n';
		expect(codes(src)).not.toContain(MISMATCH);
	});

	it('does not flag SparklineGroups.Item(1) assigned to a SparklineGroup', () => {
		const src = 'Sub S()\n    Dim rng As Range\n    Dim sg As SparklineGroup\n    Set sg = rng.SparklineGroups.Item(1)\nEnd Sub\n';
		expect(codes(src)).not.toContain(MISMATCH);
	});

	it('still indexes exactly one level (SparklineGroups(1) is a SparklineGroup; SparklineGroup.Item(1) is a Sparkline)', () => {
		const grp = 'Sub S()\n    Dim rng As Range\n    Dim sg As SparklineGroup\n    Set sg = rng.SparklineGroups(1)\nEnd Sub\n';
		const elem = 'Sub S()\n    Dim sg As SparklineGroup\n    Dim sp As Sparkline\n    Set sp = sg.Item(1)\nEnd Sub\n';
		expect(codes(grp)).not.toContain(MISMATCH);
		expect(codes(elem)).not.toContain(MISMATCH);
	});

	it('does not flag member-not-found chaining off an empty-paren call (shp.Duplicate().Group)', () => {
		const src = 'Sub S()\n    Dim shp As Shape\n    shp.Duplicate().Group\nEnd Sub\n';
		expect(codes(src)).not.toContain('member-not-found');
	});
});

describe('members typed as the type library types them (issue #90)', () => {
	// Excel's type library declares Shape.Duplicate As Shape and
	// SparklineGroup.SeriesColor As FormatColor. The Charts and Worksheets
	// properties of Application and Workbook are Sheets there; the model keeps
	// its own Charts and Worksheets for completion, and their values are Sheets.
	const CASES: ReadonlyArray<readonly [string, string]> = [
		['Dim someShape As Shape\n    Dim copied As Shape', 'Set copied = someShape.Duplicate'],
		['Dim sparkGroup As SparklineGroup\n    Dim tint As FormatColor', 'Set tint = sparkGroup.SeriesColor'],
		['Dim tabs As Sheets', 'Set tabs = Application.Charts'],
		['Dim tabs As Sheets', 'Set tabs = Application.Worksheets'],
		['Dim book As Workbook\n    Dim tabs As Sheets', 'Set tabs = book.Charts'],
		['Dim book As Workbook\n    Dim tabs As Sheets', 'Set tabs = book.Worksheets'],
	];
	it.each(CASES)('does not flag mismatch: %s ... %s', (decls, stmt) => {
		const src = `Sub S()\n    ${decls}\n    ${stmt}\nEnd Sub\n`;
		expect(codes(src)).not.toContain(MISMATCH);
	});

	it('keeps ShapeRange.Duplicate a ShapeRange', () => {
		const src = 'Sub S()\n    Dim picked As ShapeRange\n    Dim copied As Shape\n    Set copied = picked.Duplicate\nEnd Sub\n';
		expect(codes(src)).toContain(MISMATCH);
	});
});

describe('a member called with its own arguments is what it returns (issue #197)', () => {
	// Measured in Excel, Word and PowerPoint 16.0: Shapes.Range(Array(...)) is
	// a ShapeRange, GetSpellingSuggestions("x") a SpellingSuggestions and
	// SelectContentControlsByTitle("t") a ContentControls. Setting the first two
	// into the element type raises 13. Members with no parameters are indexed:
	// Placeholders(1) and GroupItems(1) are Shapes, PivotCaches(1) a
	// PivotCache. Optional parameters are still the member's own:
	// Shapes.Range(1) in PowerPoint is a ShapeRange, and
	// CommandBars.FindControls(Type:=1) a CommandBarControls.
	function hostCodes(src: string, file: string): string[] {
		return analyzeVbaModuleSource({
			source: src,
			moduleName: 'Module1',
			host: hostTokenForFileName(file),
			referencedHosts: [],
		}).diagnostics.map((d) => d.code);
	}
	const QUIET: ReadonlyArray<readonly [string, string, string]> = [
		['Book.xlsm', 'Dim sr As ShapeRange', 'Set sr = ActiveSheet.Shapes.Range(Array("A"))'],
		['Book.xlsm', 'Dim co As ChartObject', 'Set co = ActiveSheet.ChartObjects(1)'],
		['Doc.docm', 'Dim sr As ShapeRange', 'Set sr = ActiveDocument.Shapes.Range(Array("A"))'],
		['Doc.docm', 'Dim ss As SpellingSuggestions', 'Set ss = Application.GetSpellingSuggestions("helo")'],
		['Doc.docm', 'Dim ss As SpellingSuggestions', 'Set ss = Application.GetSpellingSuggestions(Word:="helo")'],
		['Doc.docm', 'Dim cc As ContentControls', 'Set cc = ActiveDocument.SelectContentControlsByTitle("t")'],
		['Deck.pptm', 'Dim sld As Slide, sr As ShapeRange', 'Set sr = sld.Shapes.Range(Array("A"))'],
		['Deck.pptm', 'Dim sld As Slide, sr As ShapeRange', 'Set sr = sld.Shapes.Range(1)'],
		['Doc.docm', 'Dim cs As CommandBarControls', 'Set cs = Application.CommandBars.FindControls(Type:=1)'],
		// A member with no parameters is indexed: Placeholders(1) and GroupItems(1) are Shapes.
		['Deck.pptm', 'Dim sld As Slide, s As Shape', 'Set s = sld.Shapes.Placeholders(1)'],
		['Book.xlsm', 'Dim g As Shape, s As Shape', 'Set s = g.GroupItems(1)'],
		['Book.xlsm', 'Dim pc As PivotCache', 'Set pc = ThisWorkbook.PivotCaches(1)'],
		// A member read off the call is read off what the call returns.
		['Book.xlsm', 'Dim sr As ShapeRange', 'Set sr = ActiveSheet.Shapes.Range(Array("A")).Duplicate'],
		['Deck.pptm', 'Dim sld As Slide, sr As ShapeRange', 'Set sr = sld.Shapes.Range(Array("A")).Duplicate'],
		['Doc.docm', 'Dim cc As ContentControl', 'Set cc = ActiveDocument.SelectContentControlsByTitle("t").Add'],
	];
	it.each(QUIET)('does not flag mismatch in %s: %s ... %s', (file, decls, stmt) => {
		const src = `Sub S()\n    ${decls}\n    ${stmt}\nEnd Sub\n`;
		expect(hostCodes(src, file)).not.toContain(MISMATCH);
	});

	const FLAGGED: ReadonlyArray<readonly [string, string, string]> = [
		['Book.xlsm', 'Dim s As Shape', 'Set s = ActiveSheet.Shapes.Range(Array("A"))'],
		['Doc.docm', 'Dim t As SpellingSuggestion', 'Set t = Application.GetSpellingSuggestions("helo")'],
		['Deck.pptm', 'Dim sld As Slide, s As Shape', 'Set s = sld.Shapes.Range(Array("A"))'],
		['Book.xlsm', 'Dim s As Shape', 'Set s = ActiveSheet.Shapes.Range(Array("A")).Duplicate'],
		['Doc.docm', 'Dim cs As ContentControls', 'Set cs = ActiveDocument.SelectContentControlsByTitle("t").Add'],
		['Deck.pptm', 'Dim sld As Slide, s As Shape', 'Set s = sld.Shapes.Range(1)'],
		['Doc.docm', 'Dim c As CommandBarControl', 'Set c = Application.CommandBars.FindControls(Type:=1)'],
		['Deck.pptm', 'Dim sld As Slide, ps As Placeholders', 'Set ps = sld.Shapes.Placeholders(1)'],
		['Book.xlsm', 'Dim g As Shape, gs As GroupShapes', 'Set gs = g.GroupItems(1)'],
		['Book.xlsm', 'Dim pcs As PivotCaches', 'Set pcs = ThisWorkbook.PivotCaches(1)'],
	];
	it.each(FLAGGED)('flags mismatch in %s: %s ... %s', (file, decls, stmt) => {
		const src = `Sub S()\n    ${decls}\n    ${stmt}\nEnd Sub\n`;
		expect(hostCodes(src, file)).toContain(MISMATCH);
	});
});
