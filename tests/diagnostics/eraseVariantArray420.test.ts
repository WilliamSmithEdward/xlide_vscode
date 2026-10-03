// Diagnostics tests: Erase on a Variant holding a dynamic array leaves it no
// storage (issue #420). Each case was run through pyVBAharness on 2026-10-02
// in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(body: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('Erase on a Variant holding Array or Split (issue #420)', () => {
	it('leaves an array with no elements', () => {
		const erased = 'Dim a As Variant\n    a = Array(10, 20, 30)\n    Erase a\n    ';
		expect(errors(`${erased}Main = UBound(a)`)).toEqual(['unallocated-dynamic-array-access']);
		expect(errors(`${erased}Main = a(0)`)).toEqual(['unallocated-dynamic-array-access']);
		expect(errors(`${erased}a(1) = 7`)).toEqual(['unallocated-dynamic-array-access']);
		expect(errors('Dim a As Variant, e As Variant\n    a = Array(10, 20, 30)\n    Erase a\n    For Each e In a\n    Next')).toEqual(['unallocated-dynamic-array-access']);
		expect(errors('Dim a As Variant\n    a = Split("x,y", ",")\n    Erase a\n    Main = UBound(a)')).toEqual(['unallocated-dynamic-array-access']);
	});

	it('stays quiet once the Variant holds an array again', () => {
		expect(errors('Dim a As Variant\n    a = Array(10, 20, 30)\n    Erase a\n    a = Array(1)\n    Main = UBound(a)')).toEqual([]);
		expect(errors('Dim a As Variant\n    a = Array(10, 20, 30)\n    Erase a\n    Main = IsArray(a)')).toEqual([]);
		expect(errors('Dim f(1 To 3) As Long, a As Variant\n    a = f\n    Main = UBound(a)')).toEqual([]);
	});

	// A Variant's copy of a fixed array is a dynamic array, which Erase
	// empties, whatever the bounds (issue #685, measured in Excel 16.0).
	it('empties a Variant that copied a fixed array', () => {
		expect(errors('Dim f(1 To 3) As Long, a As Variant\n    a = f\n    Erase a\n    Main = UBound(a)')).toEqual(['unallocated-dynamic-array-access']);
		expect(errors('Dim f(1) As Long, a As Variant\n    a = f\n    Erase a\n    Main = a(0)')).toEqual(['unallocated-dynamic-array-access']);
	});
});
