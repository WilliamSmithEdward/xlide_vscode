// Diagnostics tests: Word and PowerPoint calls whose failure the literals
// prove (issue #311). Each sample was run through pyVBAharness on
// 2026-10-02 in Word and PowerPoint 16.0.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { getPowerPointObjectModel } from '../../src/analyzer/host/powerpointObjectModel';
import { getWordObjectModel } from '../../src/analyzer/host/wordObjectModel';

const DOC = 'Dim d As Document\n    Set d = Documents.Add\n    ';
const PRES = 'Dim a As Slide, b As Slide\n    Set a = ActivePresentation.Slides.Add(1, ppLayoutBlank)\n    ';

function errors(body: string, host: 'Word' | 'PowerPoint'): string[] {
	const hostModel = host === 'Word' ? getWordObjectModel() : getPowerPointObjectModel();
	const source = `Option Explicit\nFunction Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(source, { hostModel })
		.filter((diag) => diag.code === 'host-argument-out-of-range')
		.map((diag) => /Run-time error '(-?\d+)'/.exec(diag.message)?.[1] ?? '?');
}

describe('a Word name the literal proves bad (issue #311)', () => {
	it('is refused', () => {
		for (const name of ['1bad', 'a b', 'a-b', '', 'a.b']) {
			expect(errors(`${DOC}d.Bookmarks.Add "${name}", d.Range(0, 0)`, 'Word'), name).toEqual(['5828']);
		}
		expect(errors(`${DOC}d.Styles.Add ""`, 'Word')).toEqual(['5167']);
		expect(errors(`${DOC}d.Variables.Add "zq", 1\n    d.Variables.Add "ZQ", 2`, 'Word')).toEqual(['5903']);
		expect(errors(`${DOC}d.Styles.Add "zqstyle"\n    d.Styles.Add "ZqStyle"`, 'Word')).toEqual(['5173']);
		expect(errors(`${DOC}d.CustomDocumentProperties.Add Name:="zq", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="a"\n    d.CustomDocumentProperties.Add Name:="zq", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="b"`, 'Word')).toEqual(['-2147467259']);
	});

	it('stays quiet on names Word takes', () => {
		expect(errors(`${DOC}d.Bookmarks.Add "a_b", d.Range(0, 0)`, 'Word')).toEqual([]);
		expect(errors(`${DOC}d.Bookmarks.Add "${'a'.repeat(41)}", d.Range(0, 0)`, 'Word')).toEqual([]);
		expect(errors(`${DOC}d.Variables.Add "zq", 1\n    d.Variables("zq").Delete\n    d.Variables.Add "zq", 2`, 'Word')).toEqual([]);
		expect(errors(`${DOC}d.Styles.Add "ZqStyle"\n    d.Styles.Add "ZqOther"`, 'Word')).toEqual([]);
		expect(errors(`${DOC}d.Variables.Add "zq", 1\n    Set d = Documents.Add\n    d.Variables.Add "zq", 2`, 'Word')).toEqual([]);
	});
});

describe('a PowerPoint call the literal proves bad (issue #311)', () => {
	it('is refused', () => {
		expect(errors(`${PRES}Set b = ActivePresentation.Slides.Add(2, ppLayoutBlank)\n    a.Name = "Zq"\n    b.Name = "zq"`, 'PowerPoint')).toEqual(['-2147188160']);
		expect(errors(`${PRES}a.Shapes.AddTextbox 1, 0, 0, -5, 5`, 'PowerPoint')).toEqual(['-2147024809']);
		expect(errors(`${PRES}a.Shapes.AddTextbox 1, 0, 0, 5, -5`, 'PowerPoint')).toEqual(['-2147024809']);
		expect(errors(`${PRES}a.MoveTo 0`, 'PowerPoint')).toEqual(['-2147188160']);
	});

	it('stays quiet on what PowerPoint takes', () => {
		expect(errors(`${PRES}a.Shapes.AddTextbox 1, 0, 0, 5, 5\n    a.Shapes.AddTextbox 1, 0, 0, 5, 5\n    a.Shapes(1).Name = "S"\n    a.Shapes(2).Name = "S"`, 'PowerPoint')).toEqual([]);
		expect(errors(`${PRES}a.Tags.Add "", "x"`, 'PowerPoint')).toEqual([]);
		expect(errors(`${PRES}a.Name = ""`, 'PowerPoint')).toEqual([]);
		expect(errors(`${PRES}a.MoveTo 1`, 'PowerPoint')).toEqual([]);
		expect(errors(`${PRES}Set b = a\n    a.Name = "Zq"\n    b.Name = "zq"`, 'PowerPoint')).toEqual([]);
	});
});
