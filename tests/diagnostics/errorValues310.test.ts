// Diagnostics tests: an error value read where a number or text is needed
// (issue #310). Each sample was run through pyVBAharness on 2026-10-02 in
// Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const V = 'Dim v As Variant\n    v = CVErr(xlErrNA)\n    ';

function found(body: string): string[] {
	const source = `Option Explicit\nFunction Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(source).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('an error value used as a number or text (issue #310)', () => {
	it('raises 13', () => {
		const bodies = [
			`${V}Main = v + 1`,
			`${V}Main = "a" & v`,
			`${V}Main = (v = 1)`,
			`${V}If v Then Main = 2`,
			`${V}If v Then\n        Main = 2\n    End If`,
			`${V}Dim n As Long\n    n = v`,
			`${V}Dim s As String\n    s = v`,
			`${V}Main = Val(v)`,
			`${V}Main = Abs(v)`,
			`${V}Main = Not v`,
			`${V}Select Case v\n    Case 1\n        Main = 2\n    End Select`,
			'Main = CVErr(2042) + 1',
			'Main = Len(CVErr(2042))',
			'Dim a As Variant\n    a = Array(CVErr(2007))\n    Main = a(0) * 2',
			'Main = Evaluate("1/0") + 1',
			'Range("A1").Formula = "=1/0"\n    Main = Range("A1").Value + 1',
			'Range("A1").Formula = "=NA()"\n    Dim d As Double\n    d = Range("A1").Value',
			`${V}Dim w As Variant\n    w = v\n    Main = w * 2`,
			'Dim v As Variant\n    v = "a"\n    Main = (v = CVErr(2042))',
			'Dim v As Variant\n    v = Empty\n    Main = (v = CVErr(2042))',
			`${V}Main = v & CVErr(2042)`,
		];
		for (const body of bodies) {
			expect(found(body), body).toEqual(['variant-value-misuse']);
		}
	});

	it('stays quiet where an error value is taken as it is', () => {
		const bodies = [
			`${V}Main = IsError(v)`,
			`${V}Main = TypeName(v)`,
			`${V}Dim w As Variant\n    w = v\n    Main = IsError(w)`,
			`${V}Main = CStr(v)`,
			`${V}Main = CLng(v)`,
			`${V}v = 5\n    Main = v + 1`,
			`${V}Let v = 5\n    Main = v + 1`,
			`${V}If IsError(v) Then v = 0\n    Main = v + 1`,
			'Main = IsError(Evaluate("1/0"))',
			'Range("A1").Formula = "=1/0"\n    Main = Range("A1").Text',
			'Range("A1").Formula = "=1/0"\n    Range("A1").Formula = "=1"\n    Main = Range("A1").Value + 1',
			'Dim a As Variant\n    a = Array(1, CVErr(2007))\n    Main = a(0) * 2',
			// Two error values compare.
			'Main = (CVErr(2042) = CVErr(2042))',
			`${V}Main = (v <> CVErr(2042))`,
			`${V}If v = CVErr(2042) Then Main = 2`,
			'Dim cell As Variant\n    cell = Range("A1").Value\n    If cell = CVErr(xlErrNA) Then Main = 2',
			`${V}Select Case v\n    Case CVErr(2042)\n        Main = 2\n    End Select`,
		];
		for (const body of bodies) {
			expect(found(body), body).toEqual([]);
		}
	});
});
