// Diagnostics tests: a Word document the code just added, with its text set
// by an expression, and the collections a new document leaves empty (issue
// #694). Measured on 2026-10-03 in Word 16.0 through pyVBAharness.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { getWordObjectModel } from '../../src/analyzer/host/wordObjectModel';

function found(setup: string, use: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    Dim d As Document\n    Set d = Documents.Add\n    ${setup}\n    ${use}\nEnd Function\n`;
	return analyzeModule(src, { hostModel: getWordObjectModel() })
		.filter((diag) => diag.code === 'host-argument-out-of-range')
		.map((diag) => diag.message);
}

const TWO = 'd.Content.Text = "One two." & vbCr & "Three four."';

describe('a new Word document (issue #694)', () => {
	it.each([
		['Tables(1) after text from an expression', TWO, 'Main = d.Tables(1).Rows.Count'],
		['Sections(2) after it', TWO, 'Main = d.Sections(2).Index'],
		['a bookmark after it', TWO, 'Main = d.Bookmarks("zzNope").Range.Text'],
		['Paragraphs(3) of two paragraphs', TWO, 'Main = d.Paragraphs(3).Range.Text'],
		['Paragraphs(4) of three', 'd.Content.Text = "a" & vbCr & vbCr & "b"', 'Main = d.Paragraphs(4).Range.Text'],
		['Tables(1) after text that is not known', 'd.Content.Text = Environ("COMPUTERNAME") & vbCr & "x"', 'Main = d.Tables(1).Rows.Count'],
		['Hyperlinks(1)', '', 'Main = d.Hyperlinks(1).Address'],
		['Lists(1)', '', 'Main = d.Lists(1).Range.Text'],
		['Comments(1)', '', 'Main = d.Comments(1).Range.Text'],
	])('reports %s', (_label, setup, use) => {
		const messages = found(setup, use);
		expect(messages, use).toHaveLength(1);
		expect(messages[0]).toContain("Run-time error '5941'");
	});

	it.each([
		['Paragraphs(2) of two', TWO, 'Main = d.Paragraphs(2).Range.Text'],
		['Paragraphs(2) after vbLf', 'd.Content.Text = "a" & vbLf & "b"', 'Main = d.Paragraphs(2).Range.Text'],
		['Paragraphs(2) after vbCrLf', 'd.Content.Text = "a" & vbCrLf & "b"', 'Main = d.Paragraphs(2).Range.Text'],
		['Paragraphs(2) of text not known', 'd.Content.Text = Environ("COMPUTERNAME")', 'Main = d.Paragraphs(2).Range.Text'],
		['Comments(1) after Comments.Add', 'd.Comments.Add d.Range(0, 0), "c"', 'Main = d.Comments(1).Range.Text'],
		['Tables(1) after Tables.Add', 'd.Tables.Add d.Range, 1, 1', 'Main = d.Tables(1).Rows.Count'],
		['text that reads the document', 'd.Content.Text = d.Name & vbCr', 'Main = d.Paragraphs(2).Range.Text'],
	])('stays quiet on %s', (_label, setup, use) => {
		expect(found(setup, use), use).toEqual([]);
	});
});
