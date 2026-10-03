// Diagnostics tests: a Do loop whose counter starts at its default, 0, with
// no assignment before the loop (issue #350). Measured on 2026-10-02 in
// Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(body: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

const loop = (decl: string, header: string, close = 'Loop'): string =>
	`${decl}\n    ${header}\n        a(i) = i\n        i = i + 1\n    ${close}`;

describe('a Do loop counter that starts at its default (issue #350)', () => {
	it('reports the pass past a fixed array', () => {
		for (const body of [
			loop('Dim a(2) As Long, i As Long', 'Do While i <= 3'),
			loop('Dim a(2) As Long, i As Integer', 'Do Until i > 3'),
			loop('Dim a(2) As Long, i As Long', 'While i <= 3', 'Wend'),
			loop('Dim a(2) As Long\n    Dim i As Double', 'Do While i <= 3'),
			loop('Dim a(2) As Long, i As Variant', 'Do While i <= 3'),
			loop('Dim a(2) As Long, i&', 'Do While i <= 3'),
			// Inside a block that is no loop.
			`If True Then\n    ${loop('Dim a(2) As Long, i As Long', 'Do While i <= 3')}\n    End If`,
		]) {
			expect(errors(body), body).toEqual(['array-subscript-out-of-bounds']);
		}
	});

	it('stays quiet within the bounds, after a write, or in an outer loop', () => {
		for (const body of [
			loop('Dim a(2) As Long, i As Long', 'Do While i < 3'),
			loop('Dim a(2) As Long, i As Long', 'Do Until i >= 3'),
			loop('Dim a(5) As Long, i As Long', 'Do While i <= 5'),
			loop('Dim a(2) As Long, i As Long\n    i = 1\n    Main = i', 'Do While i <= 2'),
			'Dim a(2) As Long, i As Long, k As Long\n    For k = 1 To 2\n        Do While i <= 2\n            a(i) = i\n            i = i + 1\n        Loop\n    Next',
			loop('Dim a(2) As Long, i As Long\nAgain:', 'Do While i <= 3') + '\n    If Main = 0 Then GoTo Again',
		]) {
			expect(errors(body), body).toEqual([]);
		}
	});
});
