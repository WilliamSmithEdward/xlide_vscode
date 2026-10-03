// Diagnostics tests: a guard on Len or LenB of a known String is decided
// (issue #577). Each case was run through pyVBAharness on 2026-10-02 in Excel
// 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';

// The extension's own path, which applies On Error Resume Next.
function errors(body: string): string[] {
	const source = `Option Explicit\nFunction Main() As Variant\n    Main = 1\n    ${body}\nEnd Function\n`;
	return analyzeVbaModuleSource({ source, moduleName: 'Module1', referencedHosts: [] }).diagnostics
		.filter((diag) => diag.severity === 'error').map((diag) => String(diag.code));
}

describe('a guard on Len of a known String (issue #577)', () => {
	it('skips what the guard never runs', () => {
		expect(errors('Dim a As Long, n As Long, s As String\n    If Len(s) > 1 Then n = 10 \\ a')).toEqual([]);
		expect(errors('Dim a As Long, n As Long, s As String\n    s = ""\n    If Len(s) > 1 Then\n        n = 10 \\ a\n    End If')).toEqual([]);
		expect(errors('Dim a As Long, n As Long, s As String\n    s = "ab"\n    If LenB(s) < 4 Then n = 10 \\ a')).toEqual([]);
		expect(errors('Dim a As Long, n As Long, s As String\n    If Len(s) = 0 Then Exit Function\n    n = 10 \\ a')).toEqual([]);
	});

	it('keeps the error handling the guard never changes', () => {
		expect(errors('Dim c As Collection, n As Long, s As String\n    On Error Resume Next\n    If Len(s) > 1 Then\n        On Error GoTo 0\n    End If\n    n = c.Count')).toEqual([]);
	});

	it('reports what the guard runs', () => {
		expect(errors('Dim a As Long, n As Long, s As String\n    s = "abc"\n    If Len(s) > 1 Then n = 10 \\ a')).toEqual(['division-by-zero']);
		expect(errors('Dim a As Long, n As Long, s As String\n    s = "ab"\n    If LenB(s) > 3 Then n = 10 \\ a')).toEqual(['division-by-zero']);
		expect(errors('Dim c As Collection, n As Long, s As String\n    s = "xyz"\n    On Error Resume Next\n    If Len(s) > 1 Then\n        On Error GoTo 0\n    End If\n    n = c.Count')).toEqual(['object-variable-not-set']);
	});
});
