// Diagnostics tests: a dynamic array never allocated, passed by name to the
// project's own procedure, is followed in (issue #449). Each case was run
// through pyVBAharness on 2026-10-02 in Excel 16.0 64-bit (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const HELPERS = 'Private Function FA(a() As Long) As Long\n    FA = UBound(a)\nEnd Function\n'
	+ 'Private Function FE(a() As Long) As Long\n    FE = a(0)\nEnd Function\n'
	+ 'Private Sub Fill(b() As Long)\n    ReDim b(1 To 2)\nEnd Sub\n';

function errors(body: string): string[] {
	const src = `Option Explicit\n${HELPERS}Function Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('an unallocated array passed by name (issue #449)', () => {
	it('raises 9 at its first bound or element there', () => {
		expect(errors('Dim a() As Long\n    Main = FA(a)')).toEqual(['unallocated-dynamic-array-access']);
		expect(errors('Dim a() As Long\n    Main = FE(a)')).toEqual(['unallocated-dynamic-array-access']);
	});

	it('stays quiet once allocated, for a fixed array, and where the callee ReDims it', () => {
		expect(errors('Dim a() As Long\n    ReDim a(2)\n    Main = FA(a)')).toEqual([]);
		expect(errors('Dim a(2) As Long\n    Main = FA(a)')).toEqual([]);
		expect(errors('Dim a() As Long\n    Fill a\n    Main = UBound(a)')).toEqual([]);
	});
});
