// Diagnostics tests: compile errors at a call to a project procedure (issue
// #647). Measured on 2026-10-02 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const PA = 'Private Function PA(ByVal p0 As Long, ParamArray p1() As Variant) As Long\n    PA = p0\nEnd Function\n';
const SL = 'Private Function SL(ByVal p0 As Long) As Long\n    SL = p0\nEnd Function\n';
const OC = 'Private Function OC(Optional ByVal p0 As Collection) As Long\n    OC = 1\nEnd Function\n';
const AR = 'Private Function AR(p0() As String) As Long\n    AR = 1\nEnd Function\n';
const ARD = 'Private Function ARD(p0() As Double) As Long\n    ARD = 1\nEnd Function\n';
const D = 'Dim c As New Collection, b As Boolean, d As Double, dt As Date, s As String, n As Long, a() As String\n    ';

function errors(decls: string, body: string): string[] {
	const src = `Option Explicit\n${decls}Function Main() As Variant\n    ${D}${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.message);
}

describe('compile errors at a call to a project procedure (issue #647)', () => {
	it('refuses any named argument to a procedure with a ParamArray', () => {
		for (const body of ['Main = PA(p0:=1)', 'Call PA(p0:=1)', 'Main = PA(p0:=1, p1:=2)']) {
			expect(errors(PA, body).join('\n'), body).toMatch(/Argument in ParamArray may not be named/);
		}
		expect(errors(PA, 'Main = PA(1, 2, 3)')).toEqual([]);
	});

	it('refuses a Collection passed by value to a scalar parameter', () => {
		expect(errors(SL, 'Main = SL(c)').join('\n')).toMatch(/Argument not optional/);
		expect(errors(SL, 'Dim k As Collection\n    Set k = New Collection\n    Main = SL(k)').join('\n')).toMatch(/Argument not optional/);
		expect(errors(SL, 'Main = SL(n)')).toEqual([]);
	});

	it('refuses an expression of scalars to a Collection parameter', () => {
		for (const body of ['Main = OC(b + 0)', 'Main = OC(dt + 0)', 'Main = OC(n * 2)']) {
			expect(errors(OC, body).join('\n'), body).toMatch(/Type mismatch/);
		}
		expect(errors(OC, 'Main = OC(c)')).toEqual([]);
		expect(errors(OC, 'Main = OC()')).toEqual([]);
	});

	it('refuses Empty or an expression to an array parameter', () => {
		expect(errors(AR, 'Main = AR(Empty)').join('\n')).toMatch(/array or user-defined type expected/);
		expect(errors(ARD, 'Main = ARD(d + 0)').join('\n')).toMatch(/array or user-defined type expected/);
		expect(errors(AR, 'Main = AR(s & "x")').join('\n')).toMatch(/array or user-defined type expected/);
		expect(errors(AR, 'ReDim a(1)\n    Main = AR(a)')).toEqual([]);
	});
});
