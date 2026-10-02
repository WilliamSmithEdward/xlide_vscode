// Diagnostics tests: a Type field used against its shape, and a number or
// string local given a subscript (issue #417). Each case was compiled through
// pyVBAharness on 2026-10-02 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function withField(field: string, use: string): string[] {
	return analyzeModule(
		'Option Explicit\nPrivate Type TInner\n    x As Long\nEnd Type\n'
		+ `Private Type T\n    ${field}\nEnd Type\n`
		+ 'Private Sub TakeV(ByVal v As Variant)\nEnd Sub\nPrivate Sub TakeL(ByRef n As Long)\nEnd Sub\n'
		+ `Function Main() As Variant\n    Dim t As T\n    ${use}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`,
		{ moduleKind: 'standard', moduleName: 'Module1' },
	)
		.filter((diag) => diag.severity === 'error')
		.map((diag) => diag.code);
}

function inProcedure(body: string): string[] {
	return analyzeModule(`Option Explicit\nFunction Main() As Variant\n    ${body}\nEnd Function\n`, { moduleKind: 'standard', moduleName: 'Module1' })
		.filter((diag) => diag.severity === 'error')
		.map((diag) => diag.code);
}

describe('a Type field used against its shape (issue #417)', () => {
	it('reports the compile errors the VBE gives', () => {
		expect(withField('f As Long', 'Main = t.f(1)')).toEqual(['scalar-indexed']);
		expect(withField('f As TInner', 'Main = t.f(1)')).toEqual(['scalar-indexed']);
		expect(withField('f As Long', 'Main = UBound(t.f)')).toEqual(['array-bound-requires-array']);
		expect(withField('f As Collection', 'Erase t.f')).toContain('erase-requires-array');
		expect(withField('f As String', 'Dim e As Variant\n    For Each e In t.f\n    Next')).toEqual(['for-each-source-type']);
		expect(withField('f As TInner', 'Dim e As Variant\n    For Each e In t.f\n    Next')).toEqual(['for-each-source-type']);
		expect(withField('f As Long', 'Main = (t.f Is Nothing)')).toEqual(['is-operator-non-object']);
		expect(withField('f() As Long', 't.f = 5')).toEqual(['array-target-assignment']);
		expect(withField('f(1 To 3) As Long', 'Main = t.f + 1')).toEqual(['non-scalar-binary-operand']);
		expect(withField('f() As Long', 't.f.Add 1')).toEqual(['scalar-member-access']);
		expect(withField('f As Collection', 'Main = Len(t.f)')).toEqual(['collection-operand']);
		expect(withField('f As TInner', 't.f = 5')).toEqual(['udt-value-mismatch']);
		expect(withField('f As TInner', 'TakeV t.f')).toEqual(['udt-variant-coercion']);
		expect(withField('f As TInner', 'Set t.f = New Collection')).toEqual(['set-requires-object']);
		expect(withField('f() As Long', 'TakeL t.f')).toEqual(['byref-argument-type-mismatch']);
	});

	it('stays quiet where the field fits the use', () => {
		expect(withField('f(1 To 3) As Long', 'Main = t.f(1)')).toEqual([]);
		expect(withField('f As Variant', 't.f = Array(1)\n    Main = UBound(t.f)')).toEqual([]);
		expect(withField('f As Object', 'Main = (t.f Is Nothing)')).toEqual([]);
		expect(withField('f(1 To 3) As Long', 'Dim e As Variant\n    For Each e In t.f\n    Next')).toEqual([]);
		expect(withField('f As Collection', 'Set t.f = New Collection\n    Dim e As Variant\n    For Each e In t.f\n    Next')).toEqual([]);
		expect(withField('f As Long', 'TakeL t.f')).toEqual([]);
		expect(withField('f() As Long', 'Dim a() As Long\n    ReDim a(2)\n    t.f = a')).toEqual([]);
		expect(withField('f() As Byte', 't.f = "abc"')).toEqual([]);
		expect(withField('f(1 To 3) As Long', 'Dim a() As Long\n    ReDim a(1 To 3)\n    t.f = a')).toEqual(['array-target-assignment']);
	});
});

describe('a Type field read for what its Dim left in it (issue #417)', () => {
	it('reports the run-time errors the field raises', () => {
		expect(withField('f As Object', 'Main = t.f + 1')).toEqual(['object-variable-not-set']);
		expect(withField('f As Collection', 'Main = t.f')).toEqual(['object-variable-not-set']);
		expect(withField('f As Collection', 'Set t.f = New Collection\n    Main = t.f')).toEqual(['object-default-value']);
		expect(withField('f As Collection', 'Set t.f = New Collection\n    Main = t.f(0)')).toEqual(['collection-index-out-of-range']);
		expect(withField('f As String', 'Main = t.f + 1')).toEqual(['string-arithmetic-coercion']);
		expect(withField('f As Variant', 'Main = t.f(1)')).toEqual(['variant-value-misuse']);
		expect(withField('f As Variant', 'Main = t.f.Count')).toEqual(['variant-value-misuse']);
		expect(withField('f As Variant', 'Erase t.f')).toEqual(['variant-value-misuse']);
		expect(withField('f As Variant', 'Main = (t.f Is Nothing)')).toEqual(['variant-value-misuse']);
		expect(withField('f As Collection', 'Dim e As Variant\n    For Each e In t.f\n    Next')).toEqual(['object-variable-not-set']);
		expect(withField('f() As Long', 'Dim e As Variant\n    For Each e In t.f\n    Next')).toEqual(['unallocated-dynamic-array-access']);
		expect(withField('f As String * 3', 'Main = t.f + 1')).toEqual(['string-arithmetic-coercion']);
	});

	it('stays quiet once the field is given a value', () => {
		expect(withField('f As String', 't.f = "4"\n    Main = t.f + 1')).toEqual([]);
		expect(withField('f As Variant', 't.f = Array(1, 2)\n    Main = t.f(1)')).toEqual([]);
		expect(withField('f As Collection', 'Set t.f = New Collection\n    t.f.Add 1\n    Main = t.f(1)')).toEqual([]);
		expect(withField('f As Collection', 'Set t.f = New Collection\n    With t.f\n        .Add 1\n    End With\n    Main = t.f(1)')).toEqual([]);
		expect(withField('f As String', 'Main = t.f & 1')).toEqual([]);
	});
});

describe('a local that is no array (issue #417)', () => {
	it('is reported when given a subscript or passed to UBound', () => {
		expect(inProcedure('Dim f As Long\n    Main = f(1)')).toEqual(['scalar-indexed']);
		expect(inProcedure('Dim f As String\n    Main = f(1)')).toEqual(['scalar-indexed']);
		expect(inProcedure('Dim f As Collection\n    Main = UBound(f)')).toEqual(['array-bound-requires-array']);
		expect(inProcedure('Dim f As Object\n    Main = LBound(f)')).toEqual(['array-bound-requires-array']);
	});

	it('stays quiet on a Variant, an array, a default member or the function itself', () => {
		expect(inProcedure('Dim f As Variant\n    Main = f(1)')).toEqual([]);
		expect(inProcedure('Dim f(1 To 3) As Long\n    Main = f(1)')).toEqual([]);
		expect(inProcedure('Dim f As New Collection\n    f.Add 1\n    Main = f(1)')).toEqual([]);
		expect(inProcedure('Dim f As Variant\n    Main = UBound(f)')).toEqual([]);
	});
});
