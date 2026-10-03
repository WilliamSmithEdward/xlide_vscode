// Diagnostics tests: a Nothing guard reaches a one-line If nested in a
// one-line If, and a guard decided by a value a loop never changes holds
// inside the loop (issue #575). Each case was run through pyVBAharness on
// 2026-10-02 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(body: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    Main = 1\n    ${body}\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a guard reaches a nested one-line If (issue #575)', () => {
	it('keeps the outer Nothing guard', () => {
		expect(errors('Dim c As Collection\n    If Not c Is Nothing Then If c.Count > 0 Then c.Remove 1')).toEqual([]);
		expect(errors('Dim c As Collection\n    If c Is Nothing Then Main = 0 Else If c.Count > 0 Then c.Remove 1')).toEqual([]);
	});

	it('still reports an inner access the guard does not cover', () => {
		expect(errors('Dim c As Collection\n    If c Is Nothing Then If c.Count > 0 Then c.Remove 1')).toEqual(['object-variable-not-set']);
	});
});

describe('a guard inside a loop (issue #575)', () => {
	it('holds when the loop never changes what it reads', () => {
		expect(errors('Dim c As Collection, a As Long, i As Long\n    For i = 1 To 1\n        If a > 1 Then c.Add a\n    Next')).toEqual([]);
		expect(errors('Dim c As Collection, a As Long, i As Long\n    a = 0\n    Do While i < 1\n        i = i + 1\n        If a > 1 Then\n            c.Add a\n        End If\n    Loop')).toEqual([]);
		expect(errors('Dim c As Collection, a As Long, b As Long, i As Long\n    For i = 1 To 3\n        If a > 1 Then b = 5\n        If b > 1 Then c.Add b\n    Next')).toEqual([]);
	});

	it('reports a guard a later pass turns', () => {
		expect(errors('Dim c As Collection, a As Long, b As Long, i As Long\n    For i = 1 To 3\n        If b > 1 Then c.Add b\n        If i = 2 Then b = 5\n    Next')).toEqual(['object-variable-not-set']);
	});

	it('reports when the guard runs or the loop changes its value', () => {
		expect(errors('Dim c As Collection, a As Long, i As Long\n    a = 2\n    For i = 1 To 1\n        If a > 1 Then c.Add a\n    Next')).toEqual(['object-variable-not-set']);
		expect(errors('Dim c As Collection, a As Long, i As Long\n    For i = 1 To 3\n        a = a + 1\n        If a > 1 Then c.Add a\n    Next')).toEqual(['object-variable-not-set']);
	});
});
