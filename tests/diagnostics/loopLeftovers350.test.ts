// Diagnostics tests: a For Each over a literal Split filling a Collection, and
// a ReDim whose bound is a known local, a Len or a fixed array's UBound
// (issue #350). Each case was run through pyVBAharness on 2026-10-02 in Excel
// 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(body: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a For Each over a literal Split fills a Collection (issue #350)', () => {
	it('reports a key no pass added', () => {
		expect(errors('Dim c As New Collection, p As Variant\n    For Each p In Split("a,b", ",")\n        c.Add p, p\n    Next\n    Main = c("z")')).toEqual(['collection-key-not-found']);
	});

	it('stays quiet for a key a pass added', () => {
		expect(errors('Dim c As New Collection, p As Variant\n    For Each p In Split("a,b", ",")\n        c.Add p, p\n    Next\n    Main = c("b")')).toEqual([]);
		expect(errors('Dim c As New Collection, p As Variant\n    For Each p In Array("a", "b")\n        c.Add p, p\n    Next\n    Main = c("A")')).toEqual([]);
	});
});

describe('a ReDim bound the analyzer can compute (issue #350)', () => {
	it('reads a known local, a Len and a fixed array\'s UBound', () => {
		expect(errors('Dim a() As Long, n As Long\n    n = 2 + 1\n    ReDim a(n)\n    Main = a(n + 1)')).toEqual(['array-subscript-out-of-bounds']);
		expect(errors('Dim a() As Long, s As String\n    s = "abc"\n    ReDim a(Len(s) - 1)\n    Main = a(3)')).toEqual(['array-subscript-out-of-bounds']);
		expect(errors('Dim a() As Long, b(4) As Long\n    ReDim a(UBound(b))\n    Main = a(5)')).toEqual(['array-subscript-out-of-bounds']);
	});

	it('stays quiet inside the bounds', () => {
		expect(errors('Dim a() As Long, b(4) As Long\n    ReDim a(UBound(b))\n    Main = a(4)')).toEqual([]);
		expect(errors('Dim a() As Long, n As Long\n    n = 3\n    ReDim a(n)\n    Main = a(n)')).toEqual([]);
	});
});

describe('a loop that adds zero (issue #350)', () => {
	it('leaves the divisor at zero', () => {
		expect(errors('Dim d As Long, i As Long\n    For i = 1 To 3\n        d = d + 0\n    Next\n    Main = 10 / d')).toEqual(['division-by-zero']);
		expect(errors('Dim d As Long, i As Long\n    For i = 1 To 3\n        d = d + 1\n    Next\n    Main = 10 / d')).toEqual([]);
	});
});
