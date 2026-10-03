// Diagnostics tests: what a callee leaves in the variables a call passes it
// ByRef (issue #449). Each case was run through pyVBAharness on 2026-10-02 in
// Excel 16.0 64-bit (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const HELPERS = 'Private Sub ZeroN(n As Long)\n    n = 0\nEnd Sub\n'
	+ 'Private Sub ZeroV(ByVal n As Long)\n    n = 0\nEnd Sub\n'
	+ 'Private Function ZeroF(n As Long) As Long\n    n = 0\n    ZeroF = 1\nEnd Function\n'
	+ 'Private Sub MaybeZero(n As Long, ByVal b As Boolean)\n    If b Then n = 0\nEnd Sub\n'
	+ 'Private Sub EarlyZero(n As Long)\n    If n > 3 Then Exit Sub\n    n = 0\nEnd Sub\n'
	+ 'Private Sub Free(a() As Long)\n    Erase a\nEnd Sub\n'
	+ 'Private Sub FreeRe(a() As Long)\n    Erase a\n    ReDim a(1)\nEnd Sub\n'
	+ 'Private Sub Touch(a() As Long)\n    Debug.Print UBound(a)\nEnd Sub\n';

function errors(body: string): string[] {
	const src = `Option Explicit\n${HELPERS}Function Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

const N = 'Dim n As Long, x As Long\n    n = 5\n    ';
const A = 'Dim a() As Long\n    ReDim a(2)\n    ';

describe('a value a callee leaves ByRef (issue #449)', () => {
	it('carries n = 0 back to the caller', () => {
		for (const call of ['ZeroN n', 'Call ZeroN(n)', 'x = ZeroF(n)']) {
			expect(errors(`${N}${call}\n    Main = 10 / n`), call).toEqual(['division-by-zero']);
		}
	});

	it('leaves n alone where the callee cannot or may not change it', () => {
		for (const call of ['ZeroV n', 'ZeroN (n)', 'MaybeZero n, False', 'EarlyZero n', 'ZeroN n\n    n = 2']) {
			expect(errors(`${N}${call}\n    Main = 10 / n`), call).toEqual([]);
		}
	});

	it('carries an Erase back, unless the callee allocates again', () => {
		expect(errors(`${A}Free a\n    Main = UBound(a)`)).toEqual(['unallocated-dynamic-array-access']);
		expect(errors(`${A}Call Free(a)\n    Main = a(0)`)).toEqual(['unallocated-dynamic-array-access']);
		expect(errors(`${A}FreeRe a\n    Main = UBound(a)`)).toEqual([]);
		expect(errors(`${A}Touch a\n    Main = UBound(a)`)).toEqual([]);
		expect(errors(`${A}Free a\n    ReDim a(3)\n    Main = UBound(a)`)).toEqual([]);
	});
});
