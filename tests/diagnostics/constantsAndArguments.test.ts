// Compile errors from earlier measurement batches (issue #216): constants and
// variables in each other's place, type suffixes, array arguments and named
// arguments. Each verdict is a full project compile in 64-bit Excel 16.0.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const HELPERS = 'Private Sub TakeStrArr(a() As String)\nEnd Sub\nPrivate Sub TakeVarArr(a() As Variant)\nEnd Sub\n'
	+ 'Private Function TakeLongRef(n As Long) As String\n    n = n + 1\n    TakeLongRef = CStr(n)\nEnd Function\n';
const codes = (src: string): string[] => analyzeModule(`Option Explicit\n${src}`).map((d) => d.code ?? '');
const inSub = (body: string, decls = ''): string[] => codes(`${decls}${HELPERS}Sub P1()\n${body}End Sub\n`);

describe('constants and variables in each other s place (issue #216)', () => {
	it('refuses a Const whose value is a variable, and not one from another Const', () => {
		expect(codes('Private m As Long\nPrivate Const N As Long = m\n')).toContain('const-value-not-constant');
		expect(inSub('    Dim v As Long\n    Const K As Long = v\n')).toContain('const-value-not-constant');
		expect(codes('Private Const A As Long = 1\nPrivate Const N As Long = A\n')).not.toContain('const-value-not-constant');
	});

	it('refuses a Dim bound from a variable or parameter, and not a ReDim', () => {
		expect(inSub('    Dim n As Long\n    n = 3\n    Dim a(1 To n) As Long\n')).toContain('array-bound-not-constant');
		expect(codes('Private m As Long\nPrivate b(m) As Long\n')).toContain('array-bound-not-constant');
		expect(codes('Private Sub PX(ByVal n As Long)\n    Dim a(n) As Long\nEnd Sub\n')).toContain('array-bound-not-constant');
		expect(inSub('    Dim n As Long\n    Dim a() As Long\n    n = 3\n    ReDim a(1 To n)\n')).not.toContain('array-bound-not-constant');
	});

	it('refuses a Const as a For counter or a Mid target', () => {
		expect(inSub('    For I = 1 To 2\n    Next I\n', 'Private Const I As Long = 1\n')).toContain('variable-required');
		expect(inSub('    Const J As Long = 1\n    For J = 1 To 2\n    Next J\n')).toContain('variable-required');
		expect(inSub('    Mid$(S, 1, 1) = "x"\n', 'Private Const S As String = "abc"\n')).toContain('variable-required');
		expect(inSub('    Dim s As String\n    s = "abc"\n    Mid$(s, 1, 1) = "x"\n')).not.toContain('variable-required');
	});

	it('refuses Static at module level', () => {
		expect(codes('Static m As Long\n')).toContain('static-outside-procedure');
		expect(inSub('    Static m As Long\n')).not.toContain('static-outside-procedure');
	});
});

describe('type-suffix-mismatch (issue #216)', () => {
	it.each([
		['% on a Long', '    Dim n As Long\n    n% = 2\n'],
		['$ on a Long', '    Dim n As Long\n    n$ = "x"\n'],
		['% read from a Long', '    Dim n As Long, k As Long\n    k = n% + 1\n'],
		['# on a Variant', '    Dim v As Variant\n    v# = 1\n'],
		['& after Dim n%', '    Dim n%\n    n& = 1\n'],
	])('refuses %s', (_name, body) => {
		expect(inSub(body)).toContain('type-suffix-mismatch');
	});

	it('accepts the declared type s own suffix, and a member after !', () => {
		expect(inSub('    Dim n As Long\n    n& = 1\n')).not.toContain('type-suffix-mismatch');
		expect(inSub('    Dim s As String\n    s$ = "x"\n')).not.toContain('type-suffix-mismatch');
		expect(inSub('    Dim d As Object\n    Debug.Print d!Field\n')).not.toContain('type-suffix-mismatch');
	});

	it('refuses a suffixed call through a variable that shadows the function', () => {
		expect(codes('Private Left As Long\nSub P1()\n    Debug.Print Left$("abc", 1)\nEnd Sub\n')).toContain('type-suffix-mismatch');
	});
});

describe('array arguments (issue #216)', () => {
	it('refuses a function result, or an array of another element type, for an array parameter', () => {
		expect(inSub('    TakeStrArr Split("a,b", ",")\n')).toContain('argument-shape-mismatch');
		expect(inSub('    TakeVarArr Array(1, 2)\n')).toContain('argument-shape-mismatch');
		expect(inSub('    Dim s() As String\n    TakeVarArr s\n')).toContain('argument-shape-mismatch');
	});

	it('accepts an array of the parameter s own element type', () => {
		expect(inSub('    Dim s() As String\n    TakeStrArr s\n')).not.toContain('argument-shape-mismatch');
		expect(inSub('    Dim v() As Variant\n    TakeVarArr v\n')).not.toContain('argument-shape-mismatch');
	});

	it('refuses a Variant array element passed ByRef to a Long, and not a Long element or a parenthesized one', () => {
		expect(inSub('    Dim a(1) As Variant\n    a(1) = 1\n    Debug.Print TakeLongRef(a(1))\n')).toContain('byref-argument-type-mismatch');
		expect(inSub('    Dim a(1) As Long\n    Debug.Print TakeLongRef(a(1))\n')).not.toContain('byref-argument-type-mismatch');
		expect(inSub('    Dim a(1) As Variant\n    Debug.Print TakeLongRef((a(1)))\n')).not.toContain('byref-argument-type-mismatch');
	});

	it('takes a LongLong element for a LongPtr parameter, as 64-bit Office does', () => {
		const src = 'Private Sub TakePtr(p As LongPtr)\nEnd Sub\nSub P1()\n    Dim a(1) As LongLong\n    TakePtr a(0)\nEnd Sub\n';
		expect(codes(src)).not.toContain('byref-argument-type-mismatch');
	});
});

describe('named-argument-not-allowed (issue #216)', () => {
	it.each(['InStr(String1:="abc", String2:="c")', 'Len(Expression:="abc")', 'CStr(Expression:=1)', 'CLng(Expression:=1)', 'Abs(Number:=-1)', 'Int(Number:=1.5)', 'StrComp(String1:="a", String2:="b")'])(
		'refuses %s',
		(call) => {
			expect(inSub(`    Debug.Print ${call}\n`)).toContain('named-argument-not-allowed');
		},
	);

	it.each(['Left(String:="abc", Length:=1)', 'Mid(String:="abc", Start:=1)', 'CDec(Expression:=1)', 'InStrRev(StringCheck:="abc", StringMatch:="c")', 'InStr("abc", "c")'])(
		'accepts %s',
		(call) => {
			expect(inSub(`    Debug.Print ${call}\n`)).not.toContain('named-argument-not-allowed');
		},
	);

	it('leaves a project procedure named Len alone', () => {
		expect(codes('Private Function Len(Expression As String) As Long\nEnd Function\nSub P1()\n    Debug.Print Len(Expression:="a")\nEnd Sub\n')).not.toContain('named-argument-not-allowed');
	});
});
