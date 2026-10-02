// Diagnostics tests: errors a Word document the code just added proves
// (issue #497). Each sample was measured through pyVBAharness on 2026-10-02
// in Word 16.0.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { getWordObjectModel } from '../../src/analyzer/host/wordObjectModel';
import { byCode } from '../helpers/diagnostics';

const RANGE = 'host-argument-out-of-range';
const TEXT = ['Dim d As Document', 'Set d = Documents.Add', 'd.Content.Text = "One two. Three four."'];

function hits(...lines: string[]) {
	const src = `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
	return byCode(analyzeModule(src, { hostModel: getWordObjectModel() }), RANGE);
}

describe('a new Word document (issue #497)', () => {
	it('reports members past what the document holds', () => {
		const cases: Array<[string, string]> = [
			['Main = d.Tables(1).Rows.Count', "'5941'"],
			['Main = d.Fields(1).Code.Text', "'5941'"],
			['Main = d.InlineShapes(1).Width', "'5941'"],
			['Main = d.Sections(2).Index', "'5941'"],
			['Main = d.Bookmarks("zzNope").Range.Text', "'5941'"],
			['Main = d.Paragraphs(5).Range.Text', "'5941'"],
			['Main = d.Sentences(9).Text', "'5941'"],
			['Main = d.Words(50).Text', "'5941'"],
			['Main = d.Characters(500).Text', "'5941'"],
			['Main = d.Range(0, 99999).Text', "'4608'"],
		];
		for (const [line, error] of cases) {
			const found = hits(...TEXT, line);
			expect(found, line).toHaveLength(1);
			expect(found[0].message, line).toContain(error);
		}
		expect(hits('Dim d As Document', 'Set d = Documents.Add', 'Main = d.Characters(3).Text')).toHaveLength(1);
	});

	it('stays quiet within what it holds, and once the code adds to it', () => {
		for (const lines of [
			[...TEXT, 'Main = d.Paragraphs(1).Range.Text'],
			[...TEXT, 'Main = d.Sentences(2).Text'],
			[...TEXT, 'Main = d.Characters(21).Text'],
			[...TEXT, 'd.Tables.Add d.Range(0, 0), 1, 1', 'Main = d.Tables(1).Rows.Count'],
			[...TEXT, 'd.Content.InsertAfter " Five six. Seven."', 'Main = d.Sentences(3).Text'],
			[...TEXT, 'Main = d.Range(0, 21).Text'],
			['Dim d As Document', 'Set d = Documents.Add', 'Main = d.Characters(1).Text'],
		]) {
			expect(hits(...lines), lines.join(' / ')).toHaveLength(0);
		}
	});
});
