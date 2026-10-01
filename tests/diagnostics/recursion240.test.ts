// Diagnostics tests: a Sub or Function that calls itself on every run, and
// a Function's result used past what it holds (issue #240). Each raising
// case was measured in Excel 16.0 (build 20326, 2026-09-30) and raises the
// error named every time it runs; each quiet neighbour runs clean there.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic, expectDiagnostics } from '../helpers/diagnostics';

const RECURSION = 'unbounded-recursion';

function module(main: string, rest: string): string {
	return `Option Explicit\nFunction Main() As Variant\n    ${main}\nEnd Function\n\n${rest}\n`;
}

describe('unbounded-recursion', () => {
	it.each([
		['S', 'Private Sub S()\n    S\nEnd Sub', 'S'],
		['S', 'Private Sub S()\n    Call S\nEnd Sub', 'Call S'],
		['S', 'Private Sub S()\n    Debug.Print 1\n    S\nEnd Sub', 'S'],
		['S', 'Private Sub S()\n    Dim i As Long\n    For i = 1 To 2\n        Debug.Print i\n    Next\n    S\nEnd Sub', 'S'],
		['S', 'Private Sub S()\n    With New Collection\n        .Add 1\n    End With\n    S\nEnd Sub', 'S'],
		['Main = F()', 'Private Function F() As Long\n    F = F() + 1\nEnd Function', 'F'],
		['Main = F(3)', 'Private Function F(ByVal n As Long) As Long\n    F = F(n - 1)\nEnd Function', 'F'],
		['Main = F', 'Private Function F() As Long\n    F = 1\n    F = F() * 2\nEnd Function', 'F'],
		['Main = F()', 'Private Function F() As Long\n    Call F\nEnd Function', 'Call F'],
		// A label alone changes nothing: the call still runs first.
		['S', 'Private Sub S()\nL:\n    S\nEnd Sub', 'S'],
		// With an argument it takes none of, a Variant F calls itself and indexes the result.
		['Main = F', 'Private Function F() As Variant\n    F = Array(1, 2)\n    Debug.Print F(1)\nEnd Function', 'F'],
	])('reports %s calling itself in %s', (main, rest, span) => {
		const src = module(main, rest);
		expectDiagnostic(src, analyzeModule(src), RECURSION, { span, message: ['calls itself', "'28'"] });
	});

	it('reports each procedure of a cycle, naming the chain', () => {
		const two = module('A', 'Private Sub A()\n    B\nEnd Sub\n\nPrivate Sub B()\n    A\nEnd Sub');
		expectDiagnostics(two, analyzeModule(two), RECURSION, [
			{ span: 'B', message: "'A' calls 'B', which calls 'A'," },
			{ span: 'A', message: "'B' calls 'A', which calls 'B'," },
		]);
		const three = module('A', 'Private Sub A()\n    B\nEnd Sub\n\nPrivate Sub B()\n    C\nEnd Sub\n\nPrivate Sub C()\n    A\nEnd Sub');
		const hits = byCode(analyzeModule(three), RECURSION);
		expect(hits).toHaveLength(3);
		expect(hits[0].message).toContain("'A' calls 'B', which calls 'C', which calls 'A',");
	});

	it.each([
		['a base case first', 'Main = F(3)', 'Private Function F(ByVal n As Long) As Long\n    If n <= 0 Then Exit Function\n    F = F(n - 1) + 1\nEnd Function'],
		['a call inside a loop', 'S 2', 'Private Sub S(ByVal n As Long)\n    Dim i As Long\n    For i = 1 To n\n        S 0\n    Next\nEnd Sub'],
		['the return value without parentheses', 'Main = F', 'Private Function F() As Long\n    F = 1\n    F = F * 2\nEnd Function'],
		['a single-line If', 'S 1', 'Private Sub S(ByVal n As Long)\n    If n > 1 Then S n - 1\nEnd Sub'],
		['a Function behind a single-line If', 'Main = F(1)', 'Private Function F(ByVal n As Long) As Long\n    If n > 1 Then F = F(n - 1)\nEnd Function'],
		['a Variant return value without parentheses', 'Main = F', 'Private Function F() As Variant\n    F = 1\n    F = F * 2\nEnd Function'],
		['an Exit first', 'S', 'Private Sub S()\n    Exit Sub\n    S\nEnd Sub'],
		['a block that may exit', 'S 0', 'Private Sub S(ByVal n As Long)\n    If n = 0 Then\n        Exit Sub\n    End If\n    S n\nEnd Sub'],
		['a handler first', 'S', 'Private Sub S()\n    On Error GoTo H\n    Err.Raise 5\n    S\n    Exit Sub\nH:\nEnd Sub'],
		['a callee that returns', 'A', 'Private Sub A()\n    B\nEnd Sub\n\nPrivate Sub B()\nEnd Sub'],
	])('stays quiet with %s', (_label, main, rest) => {
		expect(byCode(analyzeModule(module(main, rest)), RECURSION)).toHaveLength(0);
	});

	it('leaves a typed F(1), which does not compile, alone', () => {
		const src = module('Main = F()(0)', 'Private Function F() As Long()\n    Dim r(1) As Long\n    F = r\n    Debug.Print F(1)\nEnd Function');
		expect(byCode(analyzeModule(src), RECURSION)).toHaveLength(0);
	});

	it('leaves a Property to recursive-property-accessor', () => {
		const src = module('Main = P', 'Private Property Get P() As Long\n    P = P + 1\nEnd Property');
		expect(byCode(analyzeModule(src), RECURSION)).toHaveLength(0);
	});
});

describe('a Function whose object result is never set', () => {
	it.each([
		['Main = F().Count', 'Private Function F() As Collection\nEnd Function', 'F()'],
		['Main = F.Count', 'Private Function F() As Collection\nEnd Function', 'F'],
		['Main = F(1).Count', 'Private Function F(ByVal n As Long) As Collection\n    Debug.Print n\nEnd Function', 'F(1)'],
	])('reports %s', (main, rest, span) => {
		const src = module(main, rest);
		expectDiagnostic(src, analyzeModule(src), 'object-variable-not-set', { span, message: ['returns Nothing', "'91'"] });
	});

	it('stays quiet where the result is set, compared or no object', () => {
		const quiet = [
			module('Main = F().Count', 'Private Function F() As Collection\n    Set F = New Collection\nEnd Function'),
			module('Main = F() Is Nothing', 'Private Function F() As Collection\nEnd Function'),
			module('Main = IsEmpty(F())', 'Private Function F() As Variant\nEnd Function'),
			module('Main = F().Count', 'Private Function F() As Collection\n    Err.Raise 5\nEnd Function'),
			// A Variant returns Empty, which is no object: 424, not 91.
			module('Main = F().Count', 'Private Function F() As Variant\nEnd Function'),
			// F takes an argument, so F.Count does not compile.
			module('Main = F.Count', 'Private Function F(ByVal n As Long) As Collection\nEnd Function'),
		];
		for (const src of quiet) {
			expect(byCode(analyzeModule(src), 'object-variable-not-set'), src).toHaveLength(0);
		}
	});
});

describe('a subscript on the array a Function returns', () => {
	it.each([
		['Main = F()(5)', 'Private Function F() As Long()\n    Dim r(1) As Long\n    F = r\nEnd Function', '5'],
		['Main = F()(5)', 'Private Function F() As Variant\n    F = Array(1, 2)\nEnd Function', '5'],
		['Main = F(1)(5)', 'Private Function F(ByVal n As Long) As Variant\n    F = Array(n, n)\nEnd Function', '5'],
		['Main = F()(2)', 'Private Function F() As Variant\n    F = Split("a b")\nEnd Function', '2'],
	])('reports %s', (main, rest, span) => {
		const src = module(main, rest);
		expectDiagnostic(src, analyzeModule(src), 'array-subscript-out-of-bounds', { span, message: ["'F()' (returned by F)", "'9'"] });
	});

	it('reads Option Base for the array Array() returns', () => {
		const src = `Option Explicit\nOption Base 1\nFunction Main() As Variant\n    Main = F()(0)\nEnd Function\n\nPrivate Function F() As Variant\n    F = Array(1, 2)\nEnd Function\n`;
		expectDiagnostic(src, analyzeModule(src), 'array-subscript-out-of-bounds', { span: '0', message: 'below the lower bound 1' });
	});

	it('stays quiet in range, with two assignments, or a dynamic array', () => {
		const quiet = [
			module('Main = F(1)(1)', 'Private Function F(ByVal n As Long) As Variant\n    F = Array(n, n)\nEnd Function'),
			module('Main = F()(2)', 'Private Function F() As Variant\n    F = Array(1, 2)\n    F = Array(1, 2, 3)\nEnd Function'),
			module('Main = F()(2)', 'Private Function F() As Long()\n    Dim r() As Long\n    ReDim r(3)\n    F = r\nEnd Function'),
			module('Main = F()(2)', 'Private Function F() As Variant\n    If Len("x") = 1 Then Exit Function\n    F = Array(1, 2)\nEnd Function'),
		];
		for (const src of quiet) {
			expect(byCode(analyzeModule(src), 'array-subscript-out-of-bounds'), src).toHaveLength(0);
		}
	});
});
