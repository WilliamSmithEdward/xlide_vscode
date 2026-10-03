// Diagnostics tests: arithmetic-overflow reads a value as it reaches the
// statement, so one stored only in a branch that does not run is not there
// (issue #565). Each case was run through pyVBAharness on 2026-10-02 in
// Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(body: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

const UNKNOWN = 'Dim a As Long, b As Long\n    b = Second(Now) + 1000\n    ';

describe('a value stored in a branch (issue #565)', () => {
	it('is not taken as definite when the branch may not run', () => {
		expect(errors(`${UNKNOWN}If b > 5000 Then a = 4\n    Main = CInt(a * 10000)`)).toEqual([]);
		expect(errors(`${UNKNOWN}If b > 5000 Then a = 40000\n    Main = CInt(a)`)).toEqual([]);
		expect(errors(`${UNKNOWN}Dim n As Integer\n    If b > 5000 Then a = 40000\n    n = a`)).toEqual([]);
	});

	it('is not there when a known guard rules the branch out', () => {
		expect(errors('Dim a As Long, b As Long\n    b = 2\n    If b Mod 2 = 0 Then\n        b = -1\n    Else\n        a = 4\n    End If\n    Main = CInt(a * 10000)')).toEqual([]);
		expect(errors('Dim a As Long, b As Long\n    b = 2\n    If b = 2 Then\n        b = -1\n    Else\n        a = 4\n    End If\n    Main = CInt(a * 10000)')).toEqual([]);
		expect(errors('Dim b As Long\n    Select Case b\n    Case 1\n        b = 4\n    End Select\n    Main = CInt(b * 10000)')).toEqual([]);
	});

	it('is there when a known guard runs the branch', () => {
		expect(errors('Dim a As Long, b As Long\n    b = 1\n    Select Case b\n    Case 1\n        b = 4\n    Case Else\n        a = 1\n    End Select\n    Main = CInt(b * 10000)')).toEqual(['arithmetic-overflow']);
		expect(errors('Dim a As Long, b As Long\n    b = 2\n    If b = 2 Then a = 4\n    Main = CInt(a * 10000)')).toEqual(['arithmetic-overflow']);
		expect(errors('Dim a As Long\n    a = 4\n    Main = CInt(a * 10000)')).toEqual(['arithmetic-overflow']);
	});
});

describe('a guard over a date part (issue #565)', () => {
	it('is decided where every value in its range agrees', () => {
		expect(errors(`${UNKNOWN}If b > 5000 Then a = 1 Else a = 40000\n    Main = CInt(a)`)).toEqual(['arithmetic-overflow']);
		expect(errors('Dim a As Long, b As Long\n    b = Minute(Now) + 100\n    If b < 100 Then a = 1 Else a = 40000\n    Main = CInt(a)')).toEqual(['arithmetic-overflow']);
		expect(errors('Dim a As Long, b As Long\n    b = Month(Now) - 13\n    If b < 0 Then a = 40000 Else a = 1\n    Main = CInt(a)')).toEqual(['arithmetic-overflow']);
	});

	it('is not decided where the range holds values either way', () => {
		expect(errors('Dim a As Long, b As Long\n    b = Second(Now)\n    If b = 30 Then a = 40000 Else a = 1\n    Main = CInt(a) + 1')).toEqual([]);
		expect(errors('Dim a As Long, b As Long\n    b = Day(Now) + 1000\n    If b > 1031 Then a = 40000\n    Main = CInt(a)')).toEqual([]);
	});
});
