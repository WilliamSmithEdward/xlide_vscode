// Diagnostics tests: #673's regressions, a GoTo whose condition is decided
// and a whole-number local given a fraction. Measured on 2026-10-03 in
// Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';

const D = 'Dim a As Long, b As Long, n As Long, s As String, c As Collection\n    ';

function errors(body: string): string[] {
	const source = `Option Explicit\nFunction Main() As Variant\n    ${D}${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeVbaModuleSource({ source, moduleName: 'Module1', moduleType: 'standard' }).diagnostics
		.filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('#673: a GoTo whose condition is decided', () => {
	it('adds no way into its label when it never jumps, and skips what it jumps over', () => {
		expect(errors('If False Then GoTo Skip385\n    s = "abc"\nSkip385:\n    If Len(s) > 1 Then GoTo Skip386\n    If c Is Nothing Then\n        n = 10 \\ a\n    End If\nSkip386:')).toEqual([]);
		expect(errors('If b <> 1 Then GoTo Skip60\n    n = c.Count\nSkip60:')).toEqual([]);
		expect(errors('If a > 2 Then GoTo Skip187\n    On Error Resume Next\nSkip187:\n    If True Then GoTo Skip188\n    On Error GoTo 0\nSkip188:\n    n = 10 \\ a')).toEqual([]);
	});

	it('keeps what every real way in agrees on', () => {
		expect(errors('s = "abcdef"\n    If b <> 0 Then GoTo Skip46\n    If Len(s) > 3 Then\n        b = 5\n    End If\nSkip46:\n    Main = CInt(b * 10000)')).toEqual(['arithmetic-overflow']);
		expect(errors('If b <> 0 Then GoTo Skip47\n    a = 0\nSkip47:\n    Main = Mid("abc", a)')).toEqual(['runtime-argument-value']);
	});
});

describe('#673: a whole-number local given a fraction holds it rounded', () => {
	it('is followed into built-in arguments as the rounded value', () => {
		expect(errors('Dim i As Integer\n    i = 1.5\n    Main = StrComp("a", "b", i)')).toEqual(['runtime-argument-value']);
		expect(errors('n = -0.5\n    Main = WeekdayName(n)')).toEqual(['runtime-argument-value']);
		expect(errors('n = 0.4\n    Main = Mid("abc", n)')).toEqual(['runtime-argument-value']);
	});

	it('stays quiet where the rounded value fits', () => {
		expect(errors('n = 2.5\n    Main = Round(n, 300)')).toEqual([]);
		expect(errors('n = 0.6\n    Main = Mid("abc", n)')).toEqual([]);
	});
});
