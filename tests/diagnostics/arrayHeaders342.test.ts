// Diagnostics tests: array misses from a state-by-use matrix (issue #342).
// Each raising sample was measured through pyVBAharness on 2026-10-02 in
// Excel 16.0 (build 20326); each quiet one runs there.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode } from '../helpers/diagnostics';

const UNALLOCATED = 'unallocated-dynamic-array-access';
const PRESERVE = 'redim-preserve-dimension-change';

function wrap(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

describe('UBound of an unallocated array in a block header (issue #342)', () => {
	it('reports the header of a For, Do, While and Select Case', () => {
		const headers = [
			['For i = 0 To UBound(a): Next'],
			['For i = 1 To 3 Step UBound(a): Next'],
			['Do While i <= UBound(a): i = i + 1: Loop'],
			['While i < UBound(a): i = i + 1: Wend'],
			['Select Case UBound(a)', 'Case 0', 'End Select'],
			['ReDim a(2)', 'Erase a', 'For i = 0 To UBound(a): Next'],
		];
		for (const lines of headers) {
			expect(byCode(analyzeModule(wrap('Dim a() As Long, i As Long', ...lines)), UNALLOCATED), lines.join(' / ')).toHaveLength(1);
		}
	});

	it('stays quiet once the array is allocated, or may be by a one-line If', () => {
		const src = wrap('Dim a() As Long, i As Long', 'ReDim a(2)', 'For i = 0 To UBound(a): Next', 'Main = i');
		expect(byCode(analyzeModule(src), UNALLOCATED)).toHaveLength(0);
		const oneLine = wrap('Dim tb() As Byte, s As String, L As Long', 's = "ab"', 'L = Len(s)', 'If L > 0 Then tb = s', 'Select Case tb(0)', 'Case 97', '    Main = 2', 'End Select');
		expect(byCode(analyzeModule(oneLine), UNALLOCATED)).toHaveLength(0);
		// The branch's `=` stores the String in the Byte array; it compares nothing.
		expect(byCode(analyzeModule(oneLine), 'string-arithmetic-coercion')).toHaveLength(0);
	});
});

describe('ReDim Preserve with no lower bound after one of 1 (issue #342)', () => {
	it('reports the lower bound moving to the Option Base', () => {
		const cases = [
			['ReDim a(1 To 3)', 'ReDim Preserve a(UBound(a) + 1)'],
			['ReDim a(1 To 3)', 'ReDim Preserve a(4)'],
			['ReDim a(1 To 2, 1 To 2)', 'ReDim Preserve a(1 To 2, 3)'],
		];
		for (const lines of cases) {
			const hits = byCode(analyzeModule(wrap('Dim a() As Long', ...lines)), PRESERVE);
			expect(hits, lines.join(' / ')).toHaveLength(1);
			expect(hits[0].message).toContain('from 1 to 0');
		}
	});

	it('stays quiet when the lower bound stays', () => {
		const cases = [
			['ReDim a(1 To 3)', 'ReDim Preserve a(1 To UBound(a) + 1)'],
			['ReDim a(1 To 3)', 'ReDim Preserve a(LBound(a) To 5)'],
			['ReDim a(3)', 'ReDim Preserve a(UBound(a) + 1)'],
		];
		for (const lines of cases) {
			expect(byCode(analyzeModule(wrap('Dim a() As Long', ...lines)), PRESERVE), lines.join(' / ')).toHaveLength(0);
		}
		const based = `Option Explicit\nOption Base 1\nFunction Main() As Variant\n    Dim a() As Long\n    ReDim a(1 To 3)\n    ReDim Preserve a(4)\nEnd Function\n`;
		expect(byCode(analyzeModule(based), PRESERVE)).toHaveLength(0);
	});
});

describe('a copy of an unallocated array (issue #342)', () => {
	it('carries the missing storage to an array or a Variant', () => {
		const cases: Array<[string, string, string]> = [
			['Dim a() As Long, b() As Long', 'Main = a(0)', UNALLOCATED],
			['Dim a() As Long, b() As Long', 'Main = UBound(a)', UNALLOCATED],
			['Dim a() As Long, b() As Long', 'a(0) = 1', UNALLOCATED],
			['Dim a As Variant, b() As Long', 'Main = UBound(a)', UNALLOCATED],
			['Dim a As Variant, b() As Long', 'Main = a(0)', UNALLOCATED],
		];
		for (const [decl, use, code] of cases) {
			expect(byCode(analyzeModule(wrap(decl, 'a = b', use)), code), `${decl} / ${use}`).toHaveLength(1);
		}
	});

	it('stays quiet for a copy of an allocated array, or a Variant given a value', () => {
		const bodies = [
			['Dim a As Variant, b() As Long', 'ReDim b(1)', 'a = b', 'Main = UBound(a)'],
			['Dim a As Variant, b() As Long', 'a = b', 'a = 5', 'Main = a'],
			['Dim a() As Long, b() As Long', 'ReDim b(2)', 'a = b', 'Main = UBound(a)'],
		];
		for (const body of bodies) {
			expect(byCode(analyzeModule(wrap(...body)), UNALLOCATED), body.join(' / ')).toHaveLength(0);
		}
	});
});

describe('Join of a ReDim\'d array (issue #342)', () => {
	it('reports a Long array and a two-dimensional one once ReDim allocates them', () => {
		for (const lines of [['Dim a() As Long', 'ReDim a(2)'], ['Dim a() As String', 'ReDim a(1, 1)']]) {
			const hits = byCode(analyzeModule(wrap(...lines, 'Main = Join(a, ",")')), 'runtime-argument-value');
			expect(hits, lines.join(' / ')).toHaveLength(1);
		}
	});

	it('stays quiet on an unallocated array and on a String array', () => {
		for (const lines of [['Dim a() As Long'], ['Dim a() As String'], ['Dim a() As String', 'ReDim a(1)']]) {
			expect(byCode(analyzeModule(wrap(...lines, 'Main = Join(a, ",")')), 'runtime-argument-value'), lines.join(' / ')).toHaveLength(0);
		}
	});
});
