// Diagnostics tests: Word and PowerPoint arguments and enum values outside
// what the host takes (issue #245). Each member was swept from -100 to 999 in
// Word and PowerPoint 16.0 (build 20326, 2026-09-30); a value is reported only
// inside a run of the sweep that raised.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

function wrap(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

function errors(host: string, ...lines: string[]): string[] {
	return analyzeModule(wrap(...lines), { host }).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('Word', () => {
	it.each([
		['ActiveDocument.Paragraphs(1).Alignment = 99', '99', 'host-property-value-out-of-range', "'5148'"],
		['ActiveDocument.Paragraphs(1).LineSpacingRule = 99', '99', 'host-property-value-out-of-range', "'5148'"],
		['ActiveDocument.Content.Font.Underline = 99', '99', 'host-property-value-out-of-range', "'5843'"],
		['ActiveDocument.Content.Font.Underline = 5', '5', 'host-property-value-out-of-range', "'5843'"],
		['Selection.MoveRight Unit:=99', '99', 'host-argument-out-of-range', "'4120'"],
		['Selection.Collapse Direction:=99', '99', 'host-argument-out-of-range', "'4120'"],
		['Selection.InsertBreak Type:=99', '99', 'host-argument-out-of-range', "'9118'"],
		['ActiveDocument.Tables.Add ActiveDocument.Range, 0, 2', '0', 'host-argument-out-of-range', "'5148'"],
		['ActiveDocument.Tables.Add ActiveDocument.Range, 1, 64', '64', 'host-argument-out-of-range', "'5148'"],
	])('reports %s', (statement, span, code, error) => {
		const src = wrap(statement);
		expectDiagnostic(src, byCode(analyzeModule(src, { host: 'word' }), code), code, { span, message: error });
	});

	it('stays quiet on what Word takes', () => {
		for (const statement of [
			'ActiveDocument.Paragraphs(1).Alignment = 9',
			'ActiveDocument.Content.Font.Underline = 55',
			'ActiveDocument.Content.Font.Underline = -1',
			'Selection.MoveRight Unit:=16',
			'Selection.Collapse Direction:=0',
			'Selection.InsertBreak Type:=11',
			'ActiveDocument.Tables.Add ActiveDocument.Range, 1, 63',
		]) {
			expect(errors('word', statement), statement).toHaveLength(0);
		}
	});
});

describe('PowerPoint', () => {
	const slide = ['Dim s As Slide', 'Set s = ActivePresentation.Slides.Add(1, 12)'];

	it.each([
		[[...slide, 's.Shapes.AddShape 999, 0, 0, 10, 10'], '999', "'-2147024809'"],
		[['ActivePresentation.Slides.Add 1, 999'], '999', "'-2147024809'"],
		[[...slide, 's.Shapes.AddTable 0, 2'], '0', "'-2147188160'"],
		[[...slide, 's.Shapes.AddTable 2, 76'], '76', "'-2147188160'"],
		[['Main = ActivePresentation.Slides.Range(0).Count'], '0', "'-2147188160'"],
	])('reports %j', (lines, span, error) => {
		const src = wrap(...lines);
		expectDiagnostic(src, byCode(analyzeModule(src, { host: 'powerpoint' }), 'host-argument-out-of-range'), 'host-argument-out-of-range', { span, message: error });
	});

	it('stays quiet on what PowerPoint takes, and on an index past the slides it may have', () => {
		for (const lines of [
			[...slide, 's.Shapes.AddShape 183, 0, 0, 10, 10'],
			[...slide, 's.Shapes.AddTable 75, 75'],
			['ActivePresentation.Slides.Add 1, 27'],
			['Main = ActivePresentation.Slides.Range(3).Count'],
		]) {
			expect(errors('powerpoint', ...lines), lines.join('; ')).toHaveLength(0);
		}
	});
});
