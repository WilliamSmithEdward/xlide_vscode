import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { getWordObjectModel } from '../../src/analyzer/host/wordObjectModel';
import { getPowerPointObjectModel } from '../../src/analyzer/host/powerpointObjectModel';

const indexing = (source: string, hostModel?: ReturnType<typeof getWordObjectModel>) =>
	analyzeModule(source, { hostModel }).filter(d => d.code === 'collection-index-out-of-range' || d.code === 'host-argument-out-of-range');

describe('runtime object collection bounds', () => {
	it('does not assume a Static procedure collection starts empty on every call', () => {
		const source = 'Option Explicit\nStatic Sub ReadThenAppend()\n Dim c As New Collection\n Debug.Print c(1)\n c.Add 1\nEnd Sub\n';
		expect(indexing(source)).toEqual([]);
	});

	it('does not infer dictionary bounds for a field that a helper can mutate', () => {
		const source = 'Option Explicit\nPrivate d As Object\nSub Main()\n Set d = CreateObject("Scripting.Dictionary")\n Fill\n Debug.Print d.Keys()(0)\nEnd Sub\nPrivate Sub Fill()\n d.Add "a", 1\nEnd Sub\n';
		expect(indexing(source)).toEqual([]);
	});

	it('retains bounds proven for fresh local collections and dictionaries', () => {
		expect(indexing('Sub Main()\n Dim c As New Collection\n Debug.Print c(1)\nEnd Sub\n')).toHaveLength(1);
		expect(indexing('Sub Main()\n Dim d As Object\n Set d = CreateObject("Scripting.Dictionary")\n Debug.Print d.Keys()(0)\nEnd Sub\n')).toHaveLength(1);
	});

	it.each([
		'ActiveDocument.Tables.Add ActiveDocument.Range(0, 0), 1, 1',
		'Application.Run "PopulateDocument"',
		'DoEvents',
	])('does not infer Word collection contents after %s', change => {
		const source = `Sub Main()\n Dim d As Document\n Set d = Documents.Add\n ${change}\n Debug.Print d.Tables(1).Rows.Count\nEnd Sub\n`;
		expect(indexing(source, getWordObjectModel())).toEqual([]);
	});

	it.each(['a.MoveTo 2', 'p.Slides.Add 3, ppLayoutBlank'])('does not emit inferred PowerPoint upper bounds: %s', use => {
		const source = `Sub Main()\n Dim p As Presentation, a As Slide\n Set p = Presentations.Add\n Set a = p.Slides.Add(1, ppLayoutBlank)\n ${use}\nEnd Sub\n`;
		expect(indexing(source, getPowerPointObjectModel())).toEqual([]);
	});

	it('keeps bounds expressed using the same collection current Count', () => {
		expect(indexing('Sub Main()\n Debug.Print ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count + 1).Name\nEnd Sub\n')).toHaveLength(1);
	});

	it.each(['A0', 'XFE1', 'A1048577', 'A0:D4'])('does not infer Areas bounds for a potential defined name: %s', address => {
		expect(indexing(`Sub Main()\n Debug.Print Range("${address}").Areas(2).Address\nEnd Sub\n`)).toEqual([]);
	});
});
