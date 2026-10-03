// Diagnostics tests: guards of values, built-ins and declarations the walk
// decides, so code behind a False one is not judged (issue #691). Measured
// on 2026-10-03 in Excel 16.0 (build 20430): `Main = 10 \ z` with z never
// assigned raises 11 where the guard is True and runs where it is False.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function divides(guard: string): boolean {
	const src = 'Option Explicit\nPrivate Enum Mode\n    mA\n    mB\nEnd Enum\nFunction Main() As Variant\n'
		+ '    Dim k As Long, d As Double, v As Variant, s As String, arr(3) As Long, c As New Collection\n'
		+ '    Dim x As Date, m As Mode, xl As Long, n As Long, f As Single, sg As Single, z As Long\n'
		+ '    k = 7: d = 2.5: x = #1/2/2000#: m = mA: xl = mA: n = 2.5: f = 0.1: sg = 2.5\n'
		+ `    Main = 1\n    If ${guard} Then\n        Main = 10 \\ z\n    End If\nEnd Function\n`;
	return analyzeModule(src).some((diag) => diag.code === 'division-by-zero');
}

describe('guards the walk decides (issue #691)', () => {
	it.each([
		['a Double with a fraction', 'd > 3', 'd < 3'],
		['a Date', 'x > #1/1/2010#', 'x < #1/1/2010#'],
		['an Enum local', 'm = mB', 'm = mA'],
		['an Enum member in a Long', 'xl = mB', 'xl = mA'],
		['Abs and Sgn', 'Sgn(k) = -1', 'Sgn(k) = 1'],
		['Int', 'Int(d) = 3', 'Int(d) = 2'],
		['Round, half to even', 'Round(d) = 3', 'Round(d) = 2'],
		['Fix of a negation', 'Fix(-d) = -3', 'Fix(-d) = -2'],
		['integer division', 'k \\ 2 = 4', 'k \\ 2 = 3'],
		['a power', 'k ^ 2 = 50', 'k ^ 2 = 49'],
		['Len of CStr', 'Len(CStr(k)) > 1', 'Len(CStr(k)) = 1'],
		['IsEmpty of a Variant never assigned', 'Not IsEmpty(v)', 'IsEmpty(v)'],
		['IsArray', 'IsArray(k)', 'IsArray(arr)'],
		['TypeName', 'TypeName(k) = "String"', 'TypeName(k) = "Long"'],
		['VarType', 'VarType(k) = vbString', 'VarType(k) = vbLong'],
		['IIf', 'IIf(k > 5, False, True)', 'IIf(k > 5, True, False)'],
		['a concatenation', '"a" & k = "a8"', '"a" & k = "a7"'],
		['UBound', 'UBound(arr) > 5', 'UBound(arr) = 3'],
		['LBound', 'LBound(arr) = 1', 'LBound(arr) = 0'],
		['the Count of a new Collection', 'c.Count > 0', 'c.Count = 0'],
	])('decides %s', (_label, falseGuard, trueGuard) => {
		expect(divides(falseGuard), falseGuard).toBe(false);
		expect(divides(trueGuard), trueGuard).toBe(true);
	});

	it.each([
		'Abs(k) < 0', 'Val("12") > 20', 'IsDate("abc")', 'IsNull(v)',
	])('decides %s False', (guard) => {
		expect(divides(guard)).toBe(false);
	});

	it.each([
		'TypeName(v) = "Empty"', 'TypeName(arr) = "Long()"', 'VarType(v) = vbEmpty', 'VarType(arr) = vbArray + vbLong',
		'n = 2', 'sg > 2.4', '-2 ^ 2 = -4', '10 / 4 = 2.5',
		// A Single compares with a literal in a way the walk does not follow: True.
		'f = 0.1',
	])('leaves %s to run', (guard) => {
		expect(divides(guard)).toBe(true);
	});
});
