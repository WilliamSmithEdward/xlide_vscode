// Diagnostics tests: a ParamArray used where the VBE refuses it (issue #445).
// Each case was compiled through pyVBAharness on 2026-10-02 in Excel 16.0
// (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const HELPERS = 'Private Function G(a As Variant) As Long\n    G = 1\nEnd Function\n'
	+ 'Private Function GV(ByVal a As Variant) As Long\n    GV = 1\nEnd Function\n'
	+ 'Private Function H(b() As Variant) As Long\n    H = 1\nEnd Function\n'
	+ 'Private Function K(ParamArray q() As Variant) As Long\n    K = 1\nEnd Function\n'
	+ 'Private Sub S(a As Variant)\nEnd Sub\n';

function errors(body: string): string[] {
	const src = `Option Explicit\nPublic Function F(ParamArray p() As Variant) As Variant\n    ${body}\n    If IsEmpty(F) Then F = 1\nEnd Function\n${HELPERS}`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a ParamArray used where the VBE refuses it (issue #445)', () => {
	it('reports Invalid ParamArray use and Can\'t assign to array', () => {
		for (const body of ['F = G(p)', 'ReDim p(3)', 'Erase p', 'F = H(p)', 'Call S(p)', 'S p', 'F = G((p))']) {
			expect(errors(body), body).toEqual(['invalid-paramarray-use']);
		}
		expect(errors('Set p = Nothing')).toEqual(['array-target-assignment']);
		const bare = analyzeModule('Option Explicit\nPublic Function F(ParamArray p) As Variant\n    F = 1\nEnd Function\n').filter((diag) => diag.severity === 'error');
		expect(bare.map((diag) => diag.code)).toContain('invalid-paramarray-use');
	});

	it('stays quiet on what compiles', () => {
		for (const body of ['F = GV(p)', 'F = K(p)', 'p = Array(1)', 'F = UBound(p)', 'F = G(p(0))']) {
			expect(errors(body), body).toEqual([]);
		}
	});
});
