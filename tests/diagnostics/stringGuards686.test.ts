// Diagnostics tests: string guards decided by the module's Option Compare,
// so code behind a false one is not judged (issue #686). Measured on
// 2026-10-03 in Excel 16.0 (build 20430): `Main = 10 \ z` with z never
// assigned raises 11 where the guard is True and runs where it is False.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function divides(guard: string, text = false): boolean {
	const src = `${text ? 'Option Compare Text\n' : ''}Option Explicit\nFunction Main() As Variant\n    Dim s As String, z As Long\n    s = "a"\n    Main = 1\n    If ${guard} Then\n        Main = 10 \\ z\n    End If\nEnd Function\n`;
	return analyzeModule(src).some((diag) => diag.code === 'division-by-zero');
}

function selects(caseValue: string, text = false): boolean {
	const src = `${text ? 'Option Compare Text\n' : ''}Option Explicit\nFunction Main() As Variant\n    Dim s As String, z As Long\n    s = "a"\n    Main = 1\n    Select Case s\n    Case ${caseValue}\n        Main = 10 \\ z\n    End Select\nEnd Function\n`;
	return analyzeModule(src).some((diag) => diag.code === 'division-by-zero');
}

// Each guard with s = "a", and whether Excel ran the division, without
// Option Compare and with Option Compare Text.
const GUARDS: ReadonlyArray<readonly [string, boolean, boolean]> = [
	['s = "A"', false, true],
	['s > "b"', false, false],
	['s < "b"', true, true],
	['s >= "a"', true, true],
	['s Like "b*"', false, false],
	['s Like "A*"', false, true],
	['s Like "[a-c]"', true, true],
	['s Like "?"', true, true],
	['InStr(s, "b") > 0', false, false],
	['InStr(s, "A") > 0', false, true],
	['InStr(1, s, "A", vbBinaryCompare) > 0', false, false],
	['InStr(1, s, "A", vbTextCompare) > 0', true, true],
	['LCase(s) = "b"', false, false],
	['UCase(s) = "A"', true, true],
	['StrComp(s, "A") = 0', false, true],
	['StrComp(s, "A", vbTextCompare) = 0', true, true],
	['Replace("abc", "B", "x") = "axc"', false, true],
	['s <> "A"', true, false],
	['s = "a"', true, true],
	['s = "b"', false, false],
	['"a" < "B"', false, true],
	['"a" > "B"', true, false],
];

describe('string guards by Option Compare (issue #686)', () => {
	it.each(GUARDS)('%s: binary %s, text %s', (guard, binary, text) => {
		expect(divides(guard), `${guard} (binary)`).toBe(binary);
		expect(divides(guard, true), `${guard} (text)`).toBe(text);
	});

	it.each([
		['"b"', false, false],
		['"A"', false, true],
		['"a"', true, true],
	])('Select Case s with Case %s: binary %s, text %s', (caseValue, binary, text) => {
		expect(selects(caseValue)).toBe(binary);
		expect(selects(caseValue, true)).toBe(text);
	});
});
