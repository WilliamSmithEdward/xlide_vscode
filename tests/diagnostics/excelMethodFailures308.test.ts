// Diagnostics tests: Excel method calls whose 1004 the code proves (issue
// #308). Each sample was run through pyVBAharness on 2026-10-02 in Excel
// 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { getWordObjectModel } from '../../src/analyzer/host/wordObjectModel';

function module(body: string): string {
	return `Option Explicit\nFunction Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
}

function found(body: string): string[] {
	return analyzeModule(module(body)).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

const NEW_SHEETS = 'Dim w1 As Worksheet, w2 As Worksheet\n    Set w1 = ThisWorkbook.Worksheets.Add\n    Set w2 = ThisWorkbook.Worksheets.Add\n    ';

describe('an Excel method the code proves fails (issue #308)', () => {
	it('reports the arguments the method refuses', () => {
		const bodies = [
			'Range("A1").AutoFill Range("B1:B3")',
			'Range("A1:A3").AutoFill Range("A1:A2")',
			'Range("A1").AutoFill Range("A1")',
			'Range("A1:B1").AutoFill Destination:=Range("A1:A5")',
			'Worksheets.Add Before:=Worksheets(1), After:=Worksheets(1)',
			'Worksheets.Add Worksheets(1), Worksheets(1)',
			'ActiveSheet.Move Before:=Worksheets(1), After:=Worksheets(1)',
			'Worksheets(1).Copy Before:=Worksheets(1), After:=Worksheets(1)',
			'Range("A1:A3").Sort Key1:=Range("Z1")',
			'Range("A1:B3").Sort Key1:=Range("C1")',
			'Main = Columns(2).Resize(, 16384).Address',
			'Main = Cells(1, 2).Resize(1, 16384).Address',
			'Names.Add Name:="A1", RefersTo:="=1"',
			'Names.Add "a b", "=1"',
			'Names.Add Name:="1abc", RefersTo:="=1"',
			'ActiveWorkbook.Names.Add Name:="R1C1", RefersTo:="=1"',
			'Names.Add Name:="r1c1", RefersTo:="=1"',
			'ActiveSheet.Names.Add Name:="XFD1048576", RefersTo:="=1"',
		];
		for (const body of bodies) {
			expect(found(body), body).toEqual(['host-argument-out-of-range']);
		}
	});

	it('does not infer clipboard absence from CutCopyMode', () => {
		expect(found('Range("A1").Copy\n    Application.CutCopyMode = False\n    Range("B1").PasteSpecial xlPasteValues')).toEqual([]);
		expect(found('Application.CutCopyMode = False\n    ActiveSheet.Range("B1").PasteSpecial')).toEqual([]);
	});

	it('reports a second added sheet given the first one\'s name in another case', () => {
		expect(found(`${NEW_SHEETS}w1.Name = "Aa"\n    w2.Name = "aa"`)).toEqual(['sheet-name-invalid']);
	});
});

describe('an Excel method used as it runs (issue #308)', () => {
	it.each([
		'Worksheets("Aa").Name = "Bb"',
		'Worksheets("Aa").Delete',
		'Application.Run "RenameSheet"',
		'DoEvents',
		'If True Then Worksheets("Aa").Name = "Bb"',
		'If True Then\n    Worksheets("Aa").Delete\n    End If',
		'Dim alias As Worksheet\n    Set alias = w1\n    alias.Name = "Bb"',
	])('forgets duplicate-name facts after an unmodeled mutation: %s', change => {
		const source = module(`${NEW_SHEETS}w1.Name = "Aa"\n    ${change}\n    w2.Name = "aa"`);
		expect(analyzeModule(source).filter((diag) => diag.code === 'sheet-name-invalid')).toEqual([]);
	});

	it('does not assume a handled name assignment succeeded', () => {
		const source = module(`On Error Resume Next\n    ${NEW_SHEETS}w1.Name = "Aa"\n    On Error GoTo 0\n    w2.Name = "aa"`);
		expect(analyzeModule(source).filter((diag) => diag.code === 'sheet-name-invalid')).toEqual([]);
	});

	it('does not assume ThisWorkbook and ActiveWorkbook are the same workbook', () => {
		const source = module('Dim w1 As Worksheet, w2 As Worksheet\n    Set w1 = ThisWorkbook.Worksheets.Add\n    Set w2 = ActiveWorkbook.Worksheets.Add\n    w1.Name = "Aa"\n    w2.Name = "aa"');
		expect(analyzeModule(source).filter((diag) => diag.code === 'sheet-name-invalid')).toEqual([]);
	});

	it('stays quiet', () => {
		const bodies = [
			'Range("A1").AutoFill Range("A1:A3")',
			'Range("A1:A2").AutoFill Destination:=Range("A1:A5")',
			'Range("A1").Copy\n    Range("B1").PasteSpecial xlPasteValues\n    Application.CutCopyMode = False',
			'Application.CutCopyMode = False\n    Range("A1").Copy\n    Range("B1").PasteSpecial xlPasteValues',
			'Main = Columns(1).Resize(, 16384).Address',
			'Names.Add Name:="Zz_ok", RefersTo:="=1"',
			'Names.Add Name:="XFE1", RefersTo:="=1"',
			'Names.Add Name:="Rate.Base", RefersTo:="=1"',
			'Names.Add Name:="_x.y", RefersTo:="=1"',
			'Range("A1:A3").Sort Key1:=Range("A1")',
			'Range("A1:B3").Sort Key1:=Range("B1")',
			'Range("A1").Sort Key1:=Range("Z1")',
			'Worksheets.Add Before:=Worksheets(1)',
			`${NEW_SHEETS}w1.Name = "Qa"\n    w1.Name = "qa"\n    w2.Name = "Qb"`,
			'Dim w1 As Worksheet, w2 As Worksheet\n    Set w1 = Worksheets(1)\n    Set w2 = Worksheets(1)\n    w1.Name = "Aa"\n    w2.Name = "aa"',
		];
		for (const body of bodies) {
			expect(found(body), body).toEqual([]);
		}
	});

	it('takes a call into the module as one that may copy', () => {
		const source = `${module('Application.CutCopyMode = False\n    CopyIt\n    Range("B1").PasteSpecial xlPasteValues')}Private Sub CopyIt()\n    Range("A1").Copy\nEnd Sub\n`;
		expect(analyzeModule(source).filter((diag) => diag.code === 'paste-with-nothing-copied')).toEqual([]);
	});

	it('leaves Word alone', () => {
		const diagnostics = analyzeModule(module('Application.CutCopyMode = False\n    Selection.PasteSpecial'), { hostModel: getWordObjectModel() });
		expect(diagnostics.filter((diag) => diag.code === 'paste-with-nothing-copied')).toEqual([]);
	});
});
