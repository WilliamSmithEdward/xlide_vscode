// Diagnostics tests: a value a call passes, followed into the procedure it
// calls (issue #449). Measured in Excel 16.0 64-bit (2026-10-02) through
// pyVBAharness; each case is one Private procedure and one call from Main.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function found(callee: string, call: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    ${call}\nEnd Function\n${callee}\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => `${diag.code}: ${diag.message}`);
}

const DIVIDES = 'Private Function F(ByVal n As Long) As Variant\n    F = 10 / n\nEnd Function';

describe('a passed value the callee cannot use', () => {
	it('divides by a 0 the call passes, as a literal or a known local', () => {
		expect(found(DIVIDES, 'Main = F(0)')).toEqual([expect.stringMatching(/^division-by-zero: This call passes 0 to 'n', and 'F' divides by it in '10 \/ n'/)]);
		expect(found(DIVIDES, 'Dim z As Long\n    z = 0\n    Main = F(z)')).toEqual([expect.stringMatching(/^division-by-zero: /)]);
		expect(found(DIVIDES, 'Main = F(2) + F(0)')).toHaveLength(1);
		expect(found(DIVIDES, 'Main = F(n:=0)')).toHaveLength(1);
	});

	it('rounds a number into an integer parameter as the call does', () => {
		expect(found(DIVIDES, 'Main = F(0.4)')).toEqual([expect.stringMatching(/passes 0\.4 to 'n', which holds 0, and/)]);
		expect(found('Private Function F(ByVal n As Double) As Variant\n    F = 10 / n\nEnd Function', 'Main = F(0.4)')).toEqual([]);
	});

	it('reads a member of Nothing', () => {
		expect(found('Private Function F(c As Collection) As Variant\n    F = c.Count\nEnd Function', 'Main = F(Nothing)')).toEqual([expect.stringMatching(/^object-variable-not-set: .*'91'/)]);
		expect(found('Private Function F(c As Collection) As Variant\n    F = c(1)\nEnd Function', 'Main = F(Nothing)')).toEqual([expect.stringMatching(/^object-variable-not-set: /)]);
	});

	it('passes a bad start or length, index or operand', () => {
		expect(found('Private Function F(ByVal p As Long) As Variant\n    F = Mid("abc", p)\nEnd Function', 'Main = F(0)')).toEqual([expect.stringMatching(/^runtime-argument-value: .*'5'/)]);
		expect(found('Private Function F(ByVal p As Long) As Variant\n    F = Left("abc", p)\nEnd Function', 'Main = F(-1)')).toEqual([expect.stringMatching(/^runtime-argument-value: /)]);
		expect(found('Private Function F(ByVal i As Long) As Variant\n    Dim a(2) As Long\n    F = a(i)\nEnd Function', 'Main = F(5)')).toEqual([expect.stringMatching(/^array-subscript-out-of-bounds: .*bounds are 0 To 2/)]);
		expect(found('Private Function F(ByVal s As String) As Variant\n    F = s + 1\nEnd Function', 'Main = F("abc")')).toEqual([expect.stringMatching(/^string-arithmetic-coercion: /)]);
	});

	it('overflows the type the callee works in', () => {
		for (const body of ['F = i * 2', 'F = 2 * i']) {
			expect(found(`Private Function F(ByVal i As Integer) As Variant\n    ${body}\nEnd Function`, 'Main = F(20000)'), body).toEqual([expect.stringMatching(/^arithmetic-overflow: .*computes 40000/)]);
		}
		expect(found('Private Function F(ByVal i As Integer) As Variant\n    F = i + 1\nEnd Function', 'Main = F(32767)')).toEqual([expect.stringMatching(/^arithmetic-overflow: /)]);
	});
});

describe('a passed value the callee can use', () => {
	it('stays quiet', () => {
		expect(found(DIVIDES, 'Main = F(2) + F(5)')).toEqual([]);
		expect(found('Private Function F(ByVal n As Long) As Variant\n    If n = 0 Then Exit Function\n    F = 10 / n\nEnd Function', 'Main = F(0)')).toEqual([]);
		expect(found('Private Function F(ByVal n As Long) As Variant\n    n = 2\n    F = 10 / n\nEnd Function', 'Main = F(0)')).toEqual([]);
		expect(found('Private Function F(ByVal i As Long) As Variant\n    Dim a(2) As Long\n    F = a(i)\nEnd Function', 'Main = F(2)')).toEqual([]);
		expect(found('Private Function F(ByVal i As Integer) As Variant\n    F = i * 2\nEnd Function', 'Main = F(100)')).toEqual([]);
		expect(found('Private Function F(c As Collection) As Variant\n    If c Is Nothing Then Set c = New Collection\n    F = c.Count\nEnd Function', 'Main = F(Nothing)')).toEqual([]);
		expect(found('Private Function F(ByVal s As String) As Variant\n    F = s + 1\nEnd Function', 'Main = F("5")')).toEqual([]);
	});
});
