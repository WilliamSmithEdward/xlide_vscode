// Diagnostics tests: the false positives and the lost report of issue #556.
// Each case was run through pyVBAharness on 2026-10-02 in Excel 16.0 (build
// 20430).

import { describe, it, expect } from 'vitest';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';

function errors(top: string, body: string): string[] {
	const source = `Option Explicit\n${top}Function Main() As Variant\n    ${body}\nEnd Function\n`;
	return analyzeVbaModuleSource({ source, moduleName: 'Module1', host: 'excel' } as Parameters<typeof analyzeVbaModuleSource>[0]).diagnostics
		.filter((diag) => diag.severity === 'error')
		.map((diag) => diag.code ?? '');
}

describe('issue #556', () => {
	it('reads True stored in a Byte as 255', () => {
		expect(errors('', 'Dim a0 As Byte\n    a0 = True\n    Main = Space(a0)')).toEqual([]);
		expect(errors('Private a0 As Byte\n', 'a0 = True\n    Main = Chr(a0)')).toEqual([]);
		expect(errors('', 'Dim a0 As Integer\n    a0 = True\n    Main = Space(a0)')).toEqual(['runtime-argument-value']);
	});

	it('gives And, Or and Imp with Null the value the other side decides', () => {
		const into = (expression: string): string[] => errors('', `Dim v As Variant\n    v = Null\n    Dim r As Long\n    r = ${expression}`);
		expect(into('(CLng(40000) Or v)')).toEqual([]);
		expect(into('(v And False)')).toEqual([]);
		expect(into('(v Imp 12)')).toEqual([]);
		expect(into('(False Imp v)')).toEqual([]);
		expect(into('(v And 1)')).toEqual(['assignment-type-mismatch']);
		expect(into('(v Or 0)')).toEqual(['assignment-type-mismatch']);
		expect(into('(v Imp False)')).toEqual(['assignment-type-mismatch']);
		expect(into('(v Xor 1)')).toEqual(['assignment-type-mismatch']);
	});

	it('makes a numeric String in a Const product a Double', () => {
		expect(errors('Private Const C = &H7FFFFFFF * "2"\n', 'Main = C')).toEqual([]);
		expect(errors('Private Const C = &H7FFFFFFF * 2\n', 'Main = C')).toEqual(['const-overflow']);
	});

	it('lets CompareMode be set to the mode it already has', () => {
		const dictionary = 'Dim d As Object\n    Set d = CreateObject("Scripting.Dictionary")\n    d.Add "a", 1\n    ';
		expect(errors('', `${dictionary}d.CompareMode = 0`)).toEqual([]);
		expect(errors('', `${dictionary}d.CompareMode = 1`)).toEqual(['collection-add-argument']);
	});

	it('leaves a handler alone that no error reaches', () => {
		expect(errors('', 'Dim n As Double, d As Long\n    On Error Resume Next\n    n = 1 / d\n    On Error GoTo H2\n    Exit Function\nH2:\n    n = 1 / d')).toEqual([]);
		expect(errors('', 'Dim n As Double, d As Long\n    On Error GoTo H2\n    n = 1 / d\n    Exit Function\nH2:\n    n = 1 / d')).toEqual(['division-by-zero', 'division-by-zero']);
	});

	it('counts Item on a whole column or row by columns or rows', () => {
		expect(errors('', 'Main = ActiveSheet.Range("A1").Columns(4).EntireColumn.Item(0).Address')).toEqual([]);
		expect(errors('', 'Main = ActiveSheet.Range("A1").EntireColumn.Item(0).Address')).toEqual(['host-argument-out-of-range']);
		expect(errors('', 'Main = ActiveSheet.Range("B1").EntireRow.Item(0).Address')).toEqual(['host-argument-out-of-range']);
	});

	it('reports Null passed to an array parameter', () => {
		expect(errors('Private Sub Callee(p0() As Integer)\nEnd Sub\n', 'Callee Null')).toEqual(['argument-shape-mismatch']);
	});
});
