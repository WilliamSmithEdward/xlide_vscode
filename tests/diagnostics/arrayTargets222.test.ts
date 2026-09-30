// Diagnostics tests: whole-array assignments (issue #222). Each case was
// measured in 64-bit Excel 16.0 (build 20326, 2026-09-30).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode } from '../helpers/diagnostics';

const HELPERS = [
	'Private Function LongArr() As Long()',
	'End Function',
	'Private Function StrArr() As String()',
	'End Function',
	'Private Function VarArrT() As Variant()',
	'End Function',
	'Private Function LongArrN(ByVal n As Long) As Long()',
	'End Function',
	'Private Function VarArr() As Variant',
	'    If n0 Then VarArr = Array(1, 2) Else VarArr = Array(3)',
	'End Function',
	'Private Function Mixed() As Variant',
	'    Dim r(1) As Long',
	'    If n0 Then Mixed = Array(1) Else Mixed = r',
	'End Function',
	'Private Function Maybe() As Variant',
	'    If n0 Then Maybe = Array(1)',
	'End Function',
	'Private Function SelfRef() As Variant',
	'    If Len(SelfRef) Then SelfRef = Array(1)',
	'End Function',
	'Private Function Typed() As Long',
	'    Typed = 1',
	'End Function',
].join('\n');

function hits(...lines: string[]) {
	const src = `Option Explicit\nPrivate n0 As Boolean\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n${HELPERS}\n`;
	const diags = analyzeModule(src);
	return { compile: byCode(diags, 'array-target-assignment'), run: byCode(diags, 'assignment-type-mismatch') };
}

describe('Dim a() with no As clause is an array of Variant (issue #222)', () => {
	it.each([
		['Dim a(), b() As Long', 'a = b'],
		['Dim a(), b(1) As Long', 'a = b'],
		['Dim a(), b(1) As String', 'a = b'],
		['Dim a()', 'a = 5'],
	])('flags %s then %s at compile time', (dim, assign) => {
		expect(hits(dim, assign).compile).toHaveLength(1);
	});

	it('flags Split into it, error 13', () => {
		const { run } = hits('Dim a()', 'a = Split("a,b", ",")');
		expect(run).toHaveLength(1);
		expect(run[0].message).toContain("'13'");
	});

	it.each([
		['Dim a()', 'a = Array(1, 2)'],
		['Dim a(), b() As Variant', 'a = b'],
		['Dim a(), b()', 'a = b'],
	])('stays quiet on %s then %s', (dim, assign) => {
		const { compile, run } = hits(dim, assign);
		expect([...compile, ...run]).toHaveLength(0);
	});
});

describe('a Function declared to return a typed array (issue #222)', () => {
	it.each([
		['Dim a() As Long', 'a = StrArr()'],
		['Dim a() As String', 'a = LongArr()'],
		['Dim a() As Variant', 'a = LongArr()'],
		['Dim a() As Variant', 'a = StrArr()'],
		['Dim a()', 'a = LongArr()'],
		['Dim a()', 'a = StrArr()'],
		['Dim a() As String', 'a = LongArrN(3)'],
		['Dim a() As Integer', 'a = LongArr()'],
		['Dim a(1) As Long', 'a = LongArr()'],
	])('flags %s then %s at compile time', (dim, assign) => {
		expect(hits(dim, assign).compile).toHaveLength(1);
	});

	it.each([
		['Dim a() As Long', 'a = LongArr()'],
		['Dim a() As String', 'a = StrArr()'],
		['Dim a() As Variant', 'a = VarArrT()'],
		['Dim a()', 'a = VarArrT()'],
		['Dim a() As Long', 'a = LongArrN(3)'],
		['Dim v As Variant', 'v = LongArr()'],
	])('stays quiet on %s then %s', (dim, assign) => {
		const { compile, run } = hits(dim, assign);
		expect([...compile, ...run]).toHaveLength(0);
	});
});

describe('a Variant Function that returns only Array(...) (issue #222)', () => {
	it('flags it into an array of Long, error 13, even where it can return Empty', () => {
		expect(hits('Dim a() As Long', 'a = Maybe()').run).toHaveLength(1);
		const { run } = hits('Dim a() As Long', 'a = VarArr()');
		expect(run).toHaveLength(1);
		expect(run[0].message).toContain("'13'");
	});

	it.each([
		['Dim a() As Variant', 'a = VarArr()'],
		['Dim n As Variant', 'n = VarArr()'],
		['Dim a() As Long', 'a = Mixed()'],
		['Dim n As Long', 'n = Maybe()'],
		['Dim a() As Long', 'a = SelfRef()'],
	])('stays quiet on %s then %s', (dim, assign) => {
		const { compile, run } = hits(dim, assign);
		expect([...compile, ...run]).toHaveLength(0);
	});
});
