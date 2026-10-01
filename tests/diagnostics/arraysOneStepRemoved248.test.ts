// Diagnostics tests: arrays one step removed (issue #248). Every case was
// measured in Excel 16.0 (build 20326, 2026-09-30): the subscript count of an
// array, arrays in a user-defined type, in an array and in a Collection, Mid
// past a fixed-length string, and Len of an array.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

const TYPES = [
	'Private Type Inner',
	'    vals(3) As Long',
	'    dyn() As Long',
	'End Type',
	'Private Type Outer',
	'    arr(3) As Long',
	'    rng(1 To 3) As Long',
	'    dyn() As Long',
	'    kid As Inner',
	'    kids(1) As Inner',
	'    name As String * 3',
	'    grid(1, 2) As Long',
	'End Type',
	'Dim mg(2, 2) As Long',
	'Dim mt As Outer',
	'Dim ms As String * 3',
].join('\n');

function source(lines: readonly string[], header = '', after = ''): string {
	return `Option Explicit\n${header}${TYPES}\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n${after}`;
}

function codes(lines: readonly string[], header = '', after = ''): string[] {
	return analyzeModule(source(lines, header, after)).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

function expectReport(lines: readonly string[], code: string, span: string, message: string, header = '', after = ''): void {
	const src = source(lines, header, after);
	expectDiagnostic(src, byCode(analyzeModule(src), code), code, { span, message });
}

describe('the number of subscripts', () => {
	it.each([
		[['Dim g(2, 2) As Long', 'Main = g(1)'], 'g(1)', '2 dimensions, and 1 subscript'],
		[['Dim g(2, 2) As Long', 'Main = g(1, 1, 1)'], 'g(1, 1, 1)', '3 subscripts'],
		[['Dim g(2, 2) As Long', 'g(1) = 5'], 'g(1)', '1 subscript'],
		[['Dim g(2) As Long', 'Main = g(1, 1)'], 'g(1, 1)', '1 dimension, and 2 subscripts'],
		[['Static g(2, 2) As Long', 'Main = g(1)'], 'g(1)', '1 subscript'],
		[['Main = mg(1)'], 'mg(1)', "Array 'mg'"],
		[['Dim a(3) As Collection', 'Main = a(1, 2)'], 'a(1, 2)', '2 subscripts'],
		[['Dim t As Outer', 'Main = t.arr(1, 1)'], '(1, 1)', "Array 't.arr'"],
		[['Dim t As Outer', 'Main = t.grid(1)'], '(1)', "Array 't.grid' has 2 dimensions"],
	])('a Dim fixes the dimensions, so %j does not compile', (lines, span, message) => {
		expectReport(lines, 'wrong-number-of-dimensions', span, ['Wrong number of dimensions', message]);
	});

	it.each([
		[['Dim g() As Long', 'ReDim g(2, 2)', 'Main = g(1)'], 'g(1)', "'g' (ReDim) has 2 dimensions"],
		[['Dim g() As Long', 'ReDim g(2)', 'Main = g(1, 1, 1)'], 'g(1, 1, 1)', '3 subscripts'],
		[['Dim v As Variant', 'v = Array(1, 2)', 'Main = v(0, 0)'], 'v(0, 0)', "'v' (Array(...)) has 1 dimension"],
		[['Dim v As Variant', 'v = Array(Array(1, 2))', 'Main = v(0)(0, 0)'], '(0, 0)', "'v(0)' (Array(...))"],
	])('bounds set at run time raise 9: %j', (lines, span, message) => {
		expectReport(lines, 'array-subscript-out-of-bounds', span, ["Run-time error '9'", message]);
	});

	it('stays quiet on the right count', () => {
		expect(codes(['Dim g(2, 2) As Long', 'Main = g(1, 1)'])).toEqual([]);
		expect(codes(['Dim g() As Long', 'ReDim g(2, 2)', 'Main = g(1, 1)'])).toEqual([]);
		expect(codes(['Dim v As Variant', 'v = Array(1, 2)', 'Main = v(1)'])).toEqual([]);
		expect(codes(['Dim a(3) As Collection', 'Set a(1) = New Collection', 'a(1).Add 7', 'Main = a(1)(1)'])).toEqual([]);
	});
});

describe('arrays in a user-defined type', () => {
	it.each([
		[['Dim t As Outer', 'Main = t.arr(5)'], '5', "Subscript 5 for array 't.arr' is above the upper bound 3"],
		[['Dim t As Outer', 't.arr(-1) = 1'], '-1', 'below the lower bound 0'],
		[['Dim t As Outer', 'Main = t.kid.vals(4)'], '4', "'t.kid.vals'"],
		[['Dim t As Outer', 'Main = t.kids(2).vals(0)'], '2', "'t.kids' is above the upper bound 1"],
		[['Dim t As Outer', 'Main = t.kids(0).vals(9)'], '9', "'t.kids(0).vals'"],
		[['Dim t As Outer', 'Main = t.rng(0)'], '0', 'below the lower bound 1'],
		[['Dim t As Outer', 'Main = t.grid(1, 3)'], '3', 'in dimension 2'],
		[['Dim t As Outer', 'With t', '    Main = .arr(5)', 'End With'], '5', "'t.arr'"],
		[['Dim t As Outer', 'With t.kid', '    Main = .vals(4)', 'End With'], '4', "'t.kid.vals'"],
		[['Dim t As Outer', 'With t', '    With .kid', '        Main = .vals(4)', '    End With', 'End With'], '4', "'t.kid.vals'"],
		[['Static t As Outer', 'With t', '    Main = .grid(2, 0)', 'End With'], '2', "'t.grid'"],
		[['Dim ts(2) As Outer', 'Main = ts(1).arr(4)'], '4', "'ts(1).arr'"],
		[['Dim ts() As Outer', 'ReDim ts(2)', 'Main = ts(0).arr(9)'], '9', "'ts(0).arr'"],
		[['Static t As Outer', 'Main = t.arr(4)'], '4', "'t.arr'"],
		[['Main = mt.arr(7)'], '7', "'mt.arr'"],
	])('reports %j', (lines, span, message) => {
		expectReport(lines, 'array-subscript-out-of-bounds', span, ["Run-time error '9'", message]);
	});

	it('reads a parameter of the Type', () => {
		const after = 'Function Peek(ByRef p As Outer) As Variant\n    Peek = p.arr(5)\nEnd Function\n';
		expectReport(['Dim t As Outer', 'Main = Peek(t)'], 'array-subscript-out-of-bounds', '5', "'p.arr'", '', after);
	});

	it.each([
		[['Dim t As Outer', 'Main = UBound(t.arr, 2)'], '2', "UBound asks for dimension 2 of 't.arr', which has 1 dimension"],
		[['Dim t As Outer', 'Main = LBound(t.kid.vals, 0)'], '0', "dimension 0 of 't.kid.vals'"],
	])('reports a bound the field lacks: %j', (lines, span, message) => {
		expectReport(lines, 'array-subscript-out-of-bounds', span, message);
	});

	it('stays quiet inside the bounds, and an implicit lower bound is 0 under Option Base 1', () => {
		for (const lines of [
			['Dim t As Outer', 'Main = t.arr(3)'],
			['Dim t As Outer', 't.arr(0) = 1', 'Main = t.arr(0)'],
			['Dim t As Outer', 'Main = t.kid.vals(3)'],
			['Dim t As Outer', 'Main = t.kids(1).vals(3)'],
			['Dim t As Outer', 'Main = t.rng(1)'],
			['Dim t As Outer', 'Main = t.grid(1, 2)'],
			['Dim t As Outer', 'Main = UBound(t.grid, 2)'],
			['Dim t As Outer', 'Main = UBound(t.arr, 1)'],
			['Dim t As Outer', 'With t', '    Main = .arr(2)', 'End With'],
			['Dim ts(2) As Outer', 'Main = ts(2).arr(3)'],
		]) {
			expect(codes(lines), lines.join(': ')).toEqual([]);
		}
		expect(codes(['Dim t As Outer', 'Main = t.arr(0)'], 'Option Base 1\n')).toEqual([]);
	});

	it.each([
		[['Dim t As Outer', 'Main = t.dyn(0)'], 't.dyn(0)', "'t.dyn' has no elements here"],
		[['Dim t As Outer', 'Main = t.kid.dyn(0)'], 't.kid.dyn(0)', "'t.kid.dyn' has no elements here"],
		[['Dim t As Outer', 'ReDim t.dyn(2)', 'Erase t.dyn', 'Main = t.dyn(0)'], 't.dyn(0)', 'or that Erase emptied'],
		[['Dim t As Outer', 'Main = UBound(t.dyn)'], 't.dyn', "UBound reads the bounds of 't.dyn'"],
		[['Dim t As Outer', 'ReDim t.dyn(2)', 'Main = t.dyn(3)'], '3', "'t.dyn' (ReDim) is above the upper bound 2"],
		[['Dim t As Outer', 'ReDim t.dyn(2)', 'Main = t.dyn(1, 1)'], '(1, 1)', "'t.dyn' (ReDim) has 1 dimension"],
	])('follows a dynamic array field: %j', (lines, span, message) => {
		expectReport(lines, 'array-subscript-out-of-bounds', span, message);
	});

	it('forgets a dynamic field that anything else may fill', () => {
		const fill = 'Sub Fill(ByRef p As Outer)\n    ReDim p.dyn(5)\nEnd Sub\nSub FillArr(ByRef a() As Long)\n    ReDim a(5)\nEnd Sub\n';
		for (const lines of [
			['Dim t As Outer', 'ReDim t.dyn(2)', 'Main = t.dyn(2)'],
			['Dim t As Outer', 'Fill t', 'Main = t.dyn(0)'],
			['Dim t As Outer', 'FillArr t.dyn', 'Main = t.dyn(0)'],
			['Dim t As Outer', 'Dim u As Outer', 'ReDim u.dyn(4)', 't = u', 'Main = t.dyn(4)'],
			['Dim t As Outer', 'With t', '    ReDim .dyn(2)', 'End With', 'Main = t.dyn(1)'],
			['Dim t As Outer', 'With t', '    FillArr .dyn', '    Main = .dyn(0)', 'End With'],
			['Dim t As Outer', 'Dim i As Long', 'For i = 1 To 2', '    If i = 2 Then Main = t.dyn(0)', '    ReDim t.dyn(1)', 'Next i'],
		]) {
			expect(codes(lines, '', fill), lines.join(': ')).toEqual([]);
		}
	});
});

describe('arrays in arrays and Collections', () => {
	it.each([
		[['Dim v As Variant', 'v = Array(Array(1, 2))', 'Main = v(0)(5)'], '5', "'v(0)' (Array(...)) is above the upper bound 1"],
		[['Dim v As Variant', 'v = Array(1, Array(1, Array(1, 2)))', 'Main = v(1)(1)(2)'], '2', "'v(1)(1)'"],
		[['Dim c As New Collection', 'c.Add Array(1, 2)', 'Main = c(1)(5)'], '5', "'c(1)' (Array(...))"],
		[['Dim c As New Collection', 'c.Add 5', 'c.Add Array(1, 2)', 'Main = c(2)(2)'], '2', "'c(2)'"],
		[['Dim c As New Collection', 'c.Add Array(1, 2)', 'Main = c.Item(1)(2)'], '2', "'c(1)'"],
	])('reports %j', (lines, span, message) => {
		expectReport(lines, 'array-subscript-out-of-bounds', span, message);
	});

	it('stays quiet inside the bounds, and after a Remove reorders the items', () => {
		expect(codes(['Dim v As Variant', 'v = Array(Array(1, 2))', 'Main = v(0)(1)'])).toEqual([]);
		expect(codes(['Dim c As New Collection', 'c.Add Array(1, 2)', 'Main = c(1)(1)'])).toEqual([]);
		expect(codes(['Dim c As New Collection', 'c.Add Array(1, 2)', 'c.Remove 1', 'c.Add Array(1, 2, 3)', 'Main = c(1)(2)'])).toEqual([]);
	});
});

describe('Mid past a fixed-length string', () => {
	it.each([
		[['Dim s As String * 3', 's = "abc"', 'Mid$(s, 4, 1) = "x"'], 's, a fixed-length string of 3'],
		[['Dim s As String * 3', 'Mid(s, 4) = "x"'], 's, a fixed-length string of 3'],
		[['Dim t As Outer', 't.name = "abc"', 'Mid$(t.name, 4, 1) = "x"'], 't.name, a fixed-length string of 3'],
		[['Dim t As Outer', 'Mid$(t.name, 4, 1) = "x"'], 't.name'],
		[['Dim s(2) As String * 3', 'Mid$(s(1), 4, 1) = "x"'], 's(1)'],
		[['Const N = 3', 'Dim s As String * N', 'Mid$(s, 4, 1) = "x"'], 'a fixed-length string of 3'],
		[['Mid$(ms, 4, 1) = "x"'], 'ms'],
		[['Dim s As String', 's = "ab"', 'Mid$(s, 4, 1) = "x"'], 's, which is 2 character(s) long'],
	])('reports %j', (lines, message) => {
		expectReport([...lines, 'Main = 1'], 'runtime-argument-value', '4', message);
	});

	it('stays quiet up to the declared length, assigned or not', () => {
		expect(codes(['Dim s As String * 3', 's = "abc"', 'Mid$(s, 3, 1) = "x"', 'Main = s'])).toEqual([]);
		expect(codes(['Dim s As String * 3', 'Mid$(s, 3, 1) = "x"', 'Main = s'])).toEqual([]);
		expect(codes(['Dim t As Outer', 'Mid$(t.name, 3, 1) = "x"', 'Main = t.name'])).toEqual([]);
		expect(codes(['Dim s As String', 's = "abc"', 'Mid$(s, 3, 1) = "x"', 'Main = s'])).toEqual([]);
	});
});

describe('Len of an array', () => {
	it.each([
		[['Dim a(2) As Long', 'Main = Len(a)'], 'a', "'a' is an array, which Len"],
		[['Dim a() As Long', 'Main = Len(a)'], 'a', 'Len'],
		[['Dim a(2) As Long', 'Main = LenB(a)'], 'a', 'LenB'],
		[['Dim b(3) As Byte', 'Main = LenB(b)'], 'b', 'LenB'],
		[['Main = Len(mg)'], 'mg', "'mg'"],
		[['Dim t As Outer', 'Main = Len(t.arr)'], 't.arr', "'t.arr'"],
		[['Dim a(2) As Long', 'Main = Len((a))'], 'a', "'a'"],
	])('Len takes a variable, so %j does not compile', (lines, span, message) => {
		expectReport(lines, 'variable-required', span, ['Variable required', message]);
	});

	it('reads an array parameter', () => {
		const after = 'Function Size(ByRef Source() As Byte) As Variant\n    Size = LenB(Source)\nEnd Function\n';
		expectReport(['Main = 1'], 'variable-required', 'Source', 'LenB', '', after);
	});

	it.each([
		[['Dim a(2) As Long', 'Main = VBA.Len(a)'], 'a', 'VBA.Len cannot convert it'],
		[['Dim a() As Variant', 'Main = VBA.Len(a)'], 'a', 'not of Byte'],
		[['Dim v As Variant', 'v = Array(1)', 'Main = Len(v)'], 'v', 'which Len cannot measure'],
		[['Dim v As Variant', 'v = Array(1)', 'Main = LenB(v)'], 'v', 'which LenB cannot measure'],
		[['Main = Len(Array(1))'], 'Array(1)', 'Array(1) returns an array, which Len cannot measure'],
		[['Main = VBA.Len(Array(1))'], 'Array(1)', 'which VBA.Len cannot measure'],
		[['Main = Len(Split("a b"))'], 'Split("a b")', 'Split("a b") returns an array'],
	])('an array value raises 13: %j', (lines, span, message) => {
		expectReport(lines, 'variant-value-misuse', span, ["Run-time error '13'", message]);
	});

	it('stays quiet on a Byte array through VBA.LenB, a Type, and an element', () => {
		expect(codes(['Dim b(3) As Byte', 'Main = VBA.LenB(b)'])).toEqual([]);
		expect(codes(['Dim b() As Byte', 'b = "ab"', 'Main = VBA.LenB(b)'])).toEqual([]);
		expect(codes(['Dim t As Outer', 'Main = Len(t)'])).toEqual([]);
		expect(codes(['Dim a(2) As Long', 'Main = Len(a(1))'])).toEqual([]);
	});
});
