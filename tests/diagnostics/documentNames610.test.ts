// Diagnostics tests: #311's leftovers in Word and PowerPoint (issue #610).
// Each sample was run through pyVBAharness on 2026-10-02 in Word and
// PowerPoint 16.0, on a document or presentation the code had just added.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { getPowerPointObjectModel } from '../../src/analyzer/host/powerpointObjectModel';
import { getWordObjectModel } from '../../src/analyzer/host/wordObjectModel';

const DOC = 'Dim d As Document\n    Set d = Documents.Add\n    ';
const PRES = 'Dim p As Presentation, a As Slide, b As Slide\n    Set p = Presentations.Add(msoFalse)\n    Set a = p.Slides.Add(1, ppLayoutBlank)\n    ';

function errors(body: string, host: 'Word' | 'PowerPoint'): string[] {
	const hostModel = host === 'Word' ? getWordObjectModel() : getPowerPointObjectModel();
	const source = `Option Explicit\nFunction Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(source, { hostModel })
		.filter((diag) => diag.code === 'host-argument-out-of-range')
		.map((diag) => /Run-time error '(-?\d+)'/.exec(diag.message)?.[1] ?? '?');
}

describe('Word names, #311\'s leftovers (issue #610)', () => {
	it('reports what Word refuses', () => {
		const cases: Array<[string, string]> = [
			['With d.Bookmarks\n        .Add "a b", d.Range(0, 0)\n    End With', '5828'],
			['d.Styles.Add "Normal"', '5173'],
			['d.Styles.Add "normal"', '5173'],
			['d.Styles.Add "Heading 1"', '5173'],
			['d.Styles.Add "List Table 3 - Accent 6"', '5173'],
			['d.Styles.Add " "', '5167'],
			['d.Variables.Add "", 1', '-2147467259'],
			['d.Variables.Add "zq", 1\n    d.Variables("zq").Value = 3\n    d.Variables.Add "zq", 2', '5903'],
			['Dim d2 As Document\n    d.Variables.Add "zq", 1\n    Set d2 = d\n    d2.Variables.Add "zq", 2', '5903'],
			['d.CustomDocumentProperties.Add Name:="", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="a"', '-2147418113'],
		];
		for (const [body, error] of cases) {
			expect(errors(DOC + body, 'Word'), body).toEqual([error]);
		}
	});

	it('takes what Word takes', () => {
		for (const body of ['d.Styles.Add "Heading 10"', 'd.Styles.Add "Normalish"', 'd.Variables.Add "a b", 1', 'With d.Bookmarks\n        .Add "ab", d.Range(0, 0)\n    End With',
			'd.Variables.Add "zq", 1\n    d.Variables("zq").Delete\n    d.Variables.Add "zq", 2']) {
			expect(errors(DOC + body, 'Word'), body).toEqual([]);
		}
	});
});

describe('PowerPoint slides and shapes, #311\'s leftovers (issue #610)', () => {
	it('reports what PowerPoint refuses', () => {
		const cases: Array<[string, string]> = [
			['Set b = p.Slides.Add(2, ppLayoutBlank)\n    a.Name = b.Name', '-2147188160'],
			['Set b = p.Slides.Add(2, ppLayoutBlank)\n    a.Name = "Slide2"', '-2147188160'],
			['Set b = p.Slides.Add(2, ppLayoutBlank)\n    a.Name = "Zq"\n    p.Slides(2).Name = "zq"', '-2147188160'],
			['a.MoveTo 2', '-2147188160'],
			['p.Slides.Add 3, ppLayoutBlank', '-2147188160'],
			['a.Shapes.AddTextbox 9, 0, 0, 5, 5', '-2147024809'],
			['a.Shapes.AddTextbox 0, 0, 0, 5, 5', '-2147024809'],
			['a.Shapes.AddTextbox -2, 0, 0, 5, 5', '-2147024809'],
			['a.Shapes.AddShape 1, 0, 0, -5, 5', '-2147024809'],
			['a.Shapes.AddShape 1, 0, 0, 5, -5', '-2147024809'],
		];
		for (const [body, error] of cases) {
			expect(errors(PRES + body, 'PowerPoint'), body).toEqual([error]);
		}
	});

	it('takes what PowerPoint takes', () => {
		for (const body of ['a.Shapes.AddTextbox 1, 0, 0, 0, 0', 'a.Shapes.AddTextbox 6, 0, 0, 5, 5', 'a.Shapes.AddTextbox 1, -5, -5, 5, 5', 'a.Shapes.AddLine -5, -5, -1, -1',
			'Set b = p.Slides.Add(2, ppLayoutBlank)\n    a.MoveTo 2', 'p.Slides.Add 2, ppLayoutBlank\n    Main = a.Name', 'Set b = p.Slides.Add(1, ppLayoutBlank)\n    Main = a.Name & b.Name',
			'a.Shapes.AddShape 1, 0, 0, 0, 5']) {
			expect(errors(PRES + body, 'PowerPoint'), body).toEqual([]);
		}
	});
});
