// Diagnostics tests: non-scalar-binary-operand rule (expression-binder, v2.5.0).
//
// A bare array operand of a scalar-requiring binary operator (& concat,
// arithmetic, comparison, Boolean/bitwise) is a VBE compile error ("Type
// mismatch"). Oracle-verified rejected at compile across all four operator
// classes; scalar / Variant / indexed element a(0) / host Range operands accepted.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, spanText } from '../helpers/diagnostics';

const CODE = 'non-scalar-binary-operand';

function hits(src: string) {
	return byCode(analyzeModule(src), CODE);
}

function wrap(body: string): string {
	return `Sub T()\n${body}\nEnd Sub\n`;
}

describe('analyzeModule - non-scalar-binary-operand', () => {
	it('flags a bare fixed array as the right operand of & concatenation', () => {
		const src = wrap('    Dim a(3) As Long\n    Dim s As String\n    s = "x" & a');
		const found = hits(src);
		expect(found).toHaveLength(1);
		expect(found[0].severity).toBe('error');
		expect(spanText(src, found[0])).toBe('a');
	});

	it('flags a bare dynamic array as the left operand of & concatenation', () => {
		const src = wrap('    Dim a() As String\n    Dim s As String\n    s = a & "x"');
		const found = hits(src);
		expect(found).toHaveLength(1);
		expect(spanText(src, found[0])).toBe('a');
	});

	it('flags a bare array in an arithmetic + expression', () => {
		expect(hits(wrap('    Dim a(3) As Long\n    Dim x As Long\n    x = a + 1'))).toHaveLength(1);
	});

	it('flags a bare array as the left operand of a < comparison', () => {
		expect(hits(wrap('    Dim a(3) As Long\n    If a < 1 Then\n    End If'))).toHaveLength(1);
	});

	it('flags a bare array in an And Boolean/bitwise expression', () => {
		expect(hits(wrap('    Dim a(3) As Long\n    If a And True Then\n    End If'))).toHaveLength(1);
	});

	// --- no-false-positive controls (oracle-confirmed accepted) ---

	it('stays quiet for a scalar operand in every operator class', () => {
		expect(hits(wrap('    Dim n As Long\n    Dim s As String\n    s = "x" & n'))).toHaveLength(0);
		expect(hits(wrap('    Dim n As Long\n    Dim x As Long\n    x = n + 1'))).toHaveLength(0);
		expect(hits(wrap('    Dim n As Long\n    If n < 1 Then\n    End If'))).toHaveLength(0);
		expect(hits(wrap('    Dim n As Long\n    If n And 1 Then\n    End If'))).toHaveLength(0);
	});

	it('stays quiet for a Variant operand (may hold a scalar)', () => {
		expect(hits(wrap('    Dim v As Variant\n    Dim s As String\n    s = "x" & v'))).toHaveLength(0);
	});

	it('stays quiet for an indexed array element a(0) (a scalar)', () => {
		expect(hits(wrap('    Dim a(3) As Long\n    Dim s As String\n    s = "x" & a(0)'))).toHaveLength(0);
	});

	it('stays quiet for an Object operand (a default member may coerce)', () => {
		expect(hits(wrap('    Dim o As Object\n    Dim s As String\n    s = "x" & o'))).toHaveLength(0);
	});

	it('stays quiet for an undeclared identifier (not proven non-scalar)', () => {
		expect(hits(wrap('    Dim s As String\n    s = "x" & zzz'))).toHaveLength(0);
	});

	it('stays quiet for a UDT member access p.x (a scalar field)', () => {
		const src = 'Private Type TPoint\n    x As Long\nEnd Type\n'
			+ wrap('    Dim p As TPoint\n    Dim s As String\n    s = "x" & p.x');
		expect(hits(src)).toHaveLength(0);
	});

	it('does not apply to Is or Like (out of operator scope)', () => {
		expect(hits(wrap('    Dim a(3) As Long\n    If (a Is Nothing) Then\n    End If'))).toHaveLength(0);
		expect(hits(wrap('    Dim a(3) As Long\n    Dim b As Boolean\n    b = (a Like "x")'))).toHaveLength(0);
	});

	it('stays quiet inside an inactive conditional-compilation branch', () => {
		const src = wrap('    Dim a(3) As Long\n    Dim s As String\n#If False Then\n    s = "x" & a\n#End If');
		expect(hits(src)).toHaveLength(0);
	});

	it('flags both array operands in a chained concatenation', () => {
		// `"x" & a & b` parses as ((`"x"` & a) & b); each array operand is its own error.
		const src = wrap('    Dim a(3) As Long\n    Dim b(3) As Long\n    Dim s As String\n    s = "x" & a & b');
		expect(hits(src)).toHaveLength(2);
	});

	// --- adversarial must-stay-quiet edges (no false positives) ---

	it('stays quiet for a statement-level array-to-array copy assignment', () => {
		// `a = b` is an array-copy Assignment statement, NOT a comparison BinaryExpr,
		// so the `=` operator scope must not reach it.
		expect(hits(wrap('    Dim a() As Long\n    Dim b(3) As Long\n    a = b'))).toHaveLength(0);
	});

	it('stays quiet when an array feeds UBound and the result is the operand', () => {
		expect(hits(wrap('    Dim a(3) As Long\n    Dim n As Long\n    n = UBound(a) + 1'))).toHaveLength(0);
		expect(hits(wrap('    Dim a(3) As Long\n    If UBound(a) < 3 Then\n    End If'))).toHaveLength(0);
	});

	it('stays quiet for a parenthesised array operand (conservative)', () => {
		expect(hits(wrap('    Dim a(3) As Long\n    Dim s As String\n    s = "x" & (a)'))).toHaveLength(0);
	});
});
