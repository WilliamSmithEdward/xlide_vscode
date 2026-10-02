// Diagnostics tests: an array resized or erased while it is locked (issue
// #283). Measured in Excel 16.0 64-bit (2026-10-02) through pyVBAharness:
// each raises 10, "This array is fixed or temporarily locked".

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function found(body: string, extra = ''): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    ${body}\n    Main = 1\nEnd Function\n${extra}`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

const DYN = 'Dim a() As Long, x As Variant\n    ReDim a(2)\n    ';
const VAR = 'Dim v As Variant, x As Variant\n    v = Array(1, 2)\n    ';

describe('an array locked by For Each, With or a ByRef element', () => {
	it('cannot be erased, resized or replaced inside the loop', () => {
		for (const statement of ['Erase a', 'ReDim a(5)', 'ReDim Preserve a(5)']) {
			expect(found(`${DYN}For Each x In a\n        ${statement}\n    Next`), statement).toEqual(['array-temporarily-locked']);
		}
		expect(found(`${VAR}For Each x In v\n        Erase v\n    Next`)).toEqual(['array-temporarily-locked']);
		expect(found(`${VAR}For Each x In v\n        v = Array(9)\n    Next`)).toEqual(['array-temporarily-locked']);
	});

	it('cannot be resized inside a With on its element', () => {
		const src = 'Option Explicit\nPrivate Type T\n    n As Long\nEnd Type\nFunction Main() As Variant\n    Dim t() As T\n    ReDim t(1)\n    With t(0)\n        ReDim t(3)\n    End With\nEnd Function\n';
		expect(analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code)).toEqual(['array-temporarily-locked']);
	});

	it('cannot be erased by a callee that holds its element ByRef', () => {
		expect(found('Dim a() As Long\n    ReDim a(2)\n    Zap a(0), a', 'Private Sub Zap(ByRef e As Long, ByRef a() As Long)\n    Erase a\nEnd Sub\n')).toEqual(['array-temporarily-locked']);
	});

	it('stays quiet where nothing holds it locked, or the statement may not run', () => {
		expect(found(`${DYN}For Each x In a\n        If x > 5 Then Erase a\n    Next`)).toEqual([]);
		expect(found(`${DYN}For Each x In a\n        a(0) = 1\n    Next`)).toEqual([]);
		expect(found('Dim a() As Long, i As Long\n    ReDim a(2)\n    For i = 0 To 2\n        ReDim Preserve a(5)\n    Next')).toEqual([]);
		expect(found(`${DYN}For Each x In a\n    Next\n    Erase a`)).toEqual([]);
		expect(found('Dim a() As Long, e As Long\n    ReDim a(2)\n    Zap e, a', 'Private Sub Zap(ByRef e As Long, ByRef a() As Long)\n    Erase a\nEnd Sub\n')).toEqual([]);
	});
});
