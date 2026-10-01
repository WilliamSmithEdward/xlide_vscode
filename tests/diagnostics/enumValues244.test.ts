// Diagnostics tests: Excel enum properties and arguments given a literal
// outside what Excel takes (issue #244). Each member was swept from -100 to
// 999 in Excel 16.0 (build 20326, 2026-09-30); a value is reported only where
// the sweep saw it, or every neighbour around it, refused.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

const CODE = 'host-argument-out-of-range';
const PROPERTY = 'host-property-value-out-of-range';

function wrap(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

function codes(...lines: string[]): string[] {
	return analyzeModule(wrap(...lines)).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('enum properties', () => {
	it.each([
		['Range("A1").HorizontalAlignment', '999', '8'],
		['Range("A1").VerticalAlignment', '6', '5'],
		['Range("A1").Font.Underline', '99', '5'],
		['Range("A1").Interior.Pattern', '999', '18'],
		['Range("A1").Borders(9).LineStyle', '99', '13'],
		['Range("A1").Borders(9).Weight', '99', '4'],
		['ActiveSheet.PageSetup.Orientation', '5', '2'],
		['ActiveWindow.WindowState', '7', '3'],
	])('reports %s = %s, and runs it = %s', (target, refused, taken) => {
		const src = wrap(`${target} = ${refused}`);
		const hits = analyzeModule(src).filter((diag) => diag.severity === 'error');
		expect(hits.map((diag) => diag.code), src).toEqual([PROPERTY]);
		expect(hits[0].message).toMatch(/'(1004|9)'/);
		expect(codes(`${target} = ${taken}`), taken).toHaveLength(0);
	});

	it('leaves a constant outside the swept runs alone: xlCenter, xlPatternLinearGradient', () => {
		expect(codes('Range("A1").HorizontalAlignment = -4108')).toHaveLength(0);
		expect(codes('Range("A1").Interior.Pattern = 4000')).toHaveLength(0);
	});

	it('reports a Worksheet variable made Visible = 5', () => {
		const src = wrap('Dim ws As Worksheet', 'Set ws = Worksheets.Add', 'ws.Visible = 5');
		expectDiagnostic(src, byCode(analyzeModule(src), PROPERTY), PROPERTY, { span: '5', message: "'1004'" });
		expect(codes('Dim ws As Worksheet', 'Set ws = Worksheets.Add', 'ws.Visible = 2')).toHaveLength(0);
	});

	it('stays quiet on the properties that take any value', () => {
		for (const statement of ['Application.Calculation = 5', 'Application.CutCopyMode = 5', 'Application.ReferenceStyle = 9']) {
			expect(codes(statement), statement).toHaveLength(0);
		}
	});
});

describe('enum arguments', () => {
	it.each([
		['Main = Range("A1").Borders(99).LineStyle', '99'],
		['Main = Range("A1").End(5).Address', '5'],
		['Main = Range("A1:B2").SpecialCells(99).Address', '99'],
		['Main = Range("A1:B2").SpecialCells(13).Address', '13'],
		['Range("A1:A2").Sort Key1:=Range("A1"), Order1:=9', '9'],
		['Range("A1:A2").Sort Range("A1"), 9', '9'],
		['Range("B1").PasteSpecial Paste:=99', '99'],
		['Range("B1").PasteSpecial Paste:=9', '9'],
		['Range("A1").Insert Shift:=9', '9'],
		['Range("A1").Delete Shift:=9', '9'],
		['Range("A1").AutoFill Range("A1:A3"), 99', '99'],
	])('reports %s', (statement, span) => {
		const src = wrap(statement);
		expectDiagnostic(src, byCode(analyzeModule(src), CODE), CODE, { span, message: "'1004'" });
	});

	it.each([
		'Main = Range("A1").Borders(9).LineStyle',
		'Main = Range("A1").End(4).Address',
		'Main = Range("A1:B2").SpecialCells(2).Address',
		'Range("A1:A2").Sort Key1:=Range("A1"), Order1:=2',
		'Range("B1").PasteSpecial Paste:=12',
		'Range("A1").Insert Shift:=4',
		'Range("A1").Delete Shift:=-4159',
		'Range("A1").AutoFill Range("A1:A3"), 0',
	])('stays quiet on %s', (statement) => {
		expect(byCode(analyzeModule(wrap(statement)), CODE)).toHaveLength(0);
	});
});

