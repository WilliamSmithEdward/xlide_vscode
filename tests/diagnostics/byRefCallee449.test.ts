// Diagnostics tests: a variable passed to a procedure of the module that
// cannot change it keeps what is known about it (issue #449). Each case was
// run through pyVBAharness on 2026-10-02 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const PROCS = 'Private Sub InitV(ByVal c As Collection)\n    Set c = New Collection\nEnd Sub\n'
	+ 'Private Sub Init(c As Collection)\n    Set c = New Collection\nEnd Sub\n'
	+ 'Private Sub Touch(c As Collection)\n    If c Is Nothing Then Exit Sub\nEnd Sub\n'
	+ 'Private Sub SetN(n As Long)\n    n = 5\nEnd Sub\n'
	+ 'Private Sub ReadN(n As Long)\n    Debug.Print n\nEnd Sub\n'
	+ 'Private Sub PassOn(n As Long)\n    SetN n\nEnd Sub\n'
	+ 'Private Sub Alloc(a() As Long)\n    ReDim a(3)\nEnd Sub\n'
	+ 'Private Sub Look(a() As Long)\n    Debug.Print "x"\nEnd Sub\n';

function errors(body: string): string[] {
	const src = `Option Explicit\n${PROCS}Function Main() As Variant\n    ${body}\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a variable passed to a procedure that cannot change it (issue #449)', () => {
	it('keeps its state through a ByVal parameter, a parameter never written, and a parenthesized copy', () => {
		expect(errors('Dim c As Collection\n    InitV c\n    Main = c.Count')).toEqual(['object-variable-not-set']);
		expect(errors('Dim c As Collection\n    Touch c\n    Main = c.Count')).toEqual(['object-variable-not-set']);
		expect(errors('Dim c As Collection\n    Call Touch(c)\n    Main = c.Count')).toEqual(['object-variable-not-set']);
		expect(errors('Dim n As Long\n    SetN (n)\n    Main = 10 / n')).toEqual(['division-by-zero']);
		expect(errors('Dim n As Long\n    Call SetN((n))\n    Main = 10 / n')).toEqual(['division-by-zero']);
		expect(errors('Dim n As Long\n    ReadN n\n    Main = 10 / n')).toEqual(['division-by-zero']);
		expect(errors('Dim a() As Long\n    Look a\n    Main = UBound(a)')).toEqual(['unallocated-dynamic-array-access']);
	});

	it('forgets it where the procedure writes the parameter, or passes it on', () => {
		expect(errors('Dim c As Collection\n    Init c\n    Main = c.Count')).toEqual([]);
		expect(errors('Dim n As Long\n    SetN n\n    Main = 10 / n')).toEqual([]);
		expect(errors('Dim n As Long\n    SetN n:=n\n    Main = 10 / n')).toEqual([]);
		expect(errors('Dim n As Long\n    PassOn n\n    Main = 10 / n')).toEqual([]);
		expect(errors('Dim a() As Long\n    Alloc a\n    Main = UBound(a)')).toEqual([]);
	});
});
