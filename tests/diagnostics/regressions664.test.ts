// Diagnostics tests: #664's regressions from 09d801c (#654), each of which
// runs, and the neighbours that raise. Measured on 2026-10-03 in Excel 16.0
// (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const CALLEE = 'Private Sub Callee(ByRef p0 As Byte, ByRef p1 As Long, ByRef p2 As Variant)\nEnd Sub\n';

function errors(body: string, decls = ''): string[] {
	const src = `Option Explicit\n${decls}Function Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe("#664: the shapes that run", () => {
	it('rounds a whole-number Long with any digit count', () => {
		expect(errors('Dim a0 As Long\n    a0 = -0.5\n    Main = Round(a0, 256)')).toEqual([]);
		expect(errors('Dim a0 As Long\n    a0 = 2.5\n    Main = Round(a0, 40000)')).toEqual([]);
	});

	it('hands a Null string back from Mid without reading its Length', () => {
		expect(errors('Dim n0 As Variant, n1 As Variant, n2 As Variant\n    n0 = Null: n1 = 256: n2 = Null\n    Main = Mid(n0, n1, n2)')).toEqual([]);
		expect(errors('Dim n2 As Variant\n    n2 = Null\n    Main = Mid(Null, 1, n2)')).toEqual([]);
	});

	it('passes a Boolean holding True to a Byte as 255', () => {
		expect(errors('Dim v0 As Boolean, a1 As Long, v2 As Variant\n    v0 = True\n    Call Callee((v0), a1, v2)', CALLEE)).toEqual([]);
	});

	it('decides IsNull of a local set to Null', () => {
		expect(errors('Dim v As Variant\n    v = Null\n    If IsNull(v) Then Main = 0 Else Main = CLng(v)')).toEqual([]);
	});
});

describe('#664: the neighbours that raise', () => {
	it('still reports them', () => {
		expect(errors('Dim d As Double\n    d = 2.5\n    Main = Round(d, 256)')).toEqual(['runtime-argument-value']);
		expect(errors('Dim n0 As Variant, n2 As Variant\n    n0 = "abc": n2 = Null\n    Main = Mid(n0, 1, n2)')).toEqual(['argument-type-mismatch']);
		expect(errors('Dim v0 As Long, a1 As Long, v2 As Variant\n    v0 = 300\n    Call Callee((v0), a1, v2)', CALLEE)).toEqual(['argument-type-mismatch']);
		expect(errors('Dim v As Variant\n    v = Null\n    If Not IsNull(v) Then Main = 0 Else Main = CLng(v)')).toEqual(['argument-type-mismatch']);
	});
});
