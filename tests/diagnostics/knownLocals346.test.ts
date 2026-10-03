// Diagnostics tests: a local's known value as a For bound and as a host
// property's value (issue #346). Each case was run through pyVBAharness on
// 2026-10-02 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(body: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    ${body}\n    Main = 1\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe("a local's known value (issue #346)", () => {
	it('bounds a For loop, and is a host property value', () => {
		expect(errors('Dim v As Variant, s As Long, i As Long, t As String\n    v = Split("a,b", ",")\n    s = 1\n    For i = s To 2\n        t = t & v(i)\n    Next')).toEqual(['array-subscript-out-of-bounds']);
		expect(errors('Dim a(3) As Long, s As Long, i As Long\n    s = 2\n    For i = 0 To s + 2\n        a(i) = 1\n    Next')).toEqual(['array-subscript-out-of-bounds']);
		expect(errors('Dim n As Long\n    n = 500\n    Range("A1").Font.Size = n')).toEqual(['host-property-value-out-of-range']);
	});

	it('stays quiet where the value fits, or the loop writes the bound', () => {
		expect(errors('Dim v As Variant, s As Long, i As Long, t As String\n    v = Split("a,b", ",")\n    s = 0\n    For i = s To 1\n        t = t & v(i)\n    Next')).toEqual([]);
		expect(errors('Dim n As Long\n    n = 12\n    Range("A1").Font.Size = n')).toEqual([]);
	});
});
