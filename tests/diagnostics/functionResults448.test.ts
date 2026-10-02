// Diagnostics tests: a Function's known result, followed into its caller
// (issue #448). Measured in Excel 16.0 64-bit (2026-10-02) through
// pyVBAharness; each case is one Private Function F and one use from Main.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function found(type: string, body: string, use: string): string[] {
	const fn = `Private Function F() As ${type}\n${body ? `    ${body}\n` : ''}End Function`;
	const src = `Option Explicit\nFunction Main() As Variant\n    ${use}\nEnd Function\n${fn}\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => `${diag.code}: ${diag.message}`);
}

describe('a Function whose result its text fixes', () => {
	it('stores what it returns: a string into a number, too large a number, Null', () => {
		expect(found('String', 'F = "abc"', 'Dim n As Integer\n    n = F()\n    Main = n')).toEqual([expect.stringMatching(/^assignment-type-mismatch: .*'F\(\)', which returns "abc"/)]);
		expect(found('String', 'F = "abc"', 'Dim n As Integer\n    n = F\n    Main = n')).toEqual([expect.stringMatching(/^assignment-type-mismatch: /)]);
		expect(found('String', '', 'Dim n As Long\n    n = F()\n    Main = n')).toEqual([expect.stringMatching(/^assignment-type-mismatch: .*returns ""/)]);
		expect(found('Long', 'F = 300', 'Dim b As Byte\n    b = F()\n    Main = b')).toEqual([expect.stringMatching(/^arithmetic-overflow: .*stores 300 in a Byte/)]);
		expect(found('Variant', 'F = Null', 'Dim n As Long\n    n = F()\n    Main = n')).toEqual([expect.stringMatching(/^assignment-type-mismatch: .*'F\(\)' returns Null/)]);
	});

	it('divides by 0 and passes 0 to Mid', () => {
		expect(found('Long', '', 'Main = 10 / F()')).toEqual([expect.stringMatching(/^division-by-zero: /)]);
		expect(found('Long', '', 'Main = 10 / F')).toEqual([expect.stringMatching(/^division-by-zero: /)]);
		expect(found('Long', 'F = 0', 'Main = 10 / F()')).toEqual([expect.stringMatching(/^division-by-zero: /)]);
		expect(found('Long', '', 'Main = Mid("abc", F())')).toEqual([expect.stringMatching(/^runtime-argument-value: .*'Start' of 'Mid' is 0/)]);
	});

	it('returns Nothing or an array with no storage', () => {
		expect(found('Collection', 'Set F = Nothing', 'Main = F().Count')).toEqual([expect.stringMatching(/^object-variable-not-set: Function 'F' sets its result to Nothing/)]);
		expect(found('Long()', '', 'Main = UBound(F())')).toEqual([expect.stringMatching(/^unallocated-dynamic-array-access: .*'9'/)]);
	});
});

describe('a Function whose result is not fixed, or is fine', () => {
	it('stays quiet', () => {
		expect(found('Variant', '', 'Dim n As Long\n    n = F()\n    Main = n')).toEqual([]);
		expect(found('Long', 'F = 2', 'Main = 10 / F()')).toEqual([]);
		expect(found('String', 'F = "12"', 'Dim n As Integer\n    n = F()\n    Main = n')).toEqual([]);
		expect(found('Long', 'F = 200', 'Dim b As Byte\n    b = F()\n    Main = b')).toEqual([]);
		expect(found('Long', 'If Timer < 0 Then Exit Function\n    F = 2', 'Main = 10 / F()')).toEqual([]);
		expect(found('Long', 'F = 2\n    If Timer < 0 Then F = 0', 'Main = 10 / F()')).toEqual([]);
		expect(found('Collection', 'Set F = New Collection', 'Main = F().Count')).toEqual([]);
		expect(found('Long', 'Static k As Long\n    k = k + 1\n    F = k', 'Main = 10 / F()')).toEqual([]);
		expect(found('Long', 'F = 2', 'Main = Mid("abc", F())')).toEqual([]);
		expect(found('Long()', 'ReDim F(2)', 'Main = UBound(F())')).toEqual([]);
	});

	it('leaves a local of the same name alone', () => {
		expect(found('Long', '', 'Dim F As Long\n    F = 2\n    Main = 10 / F')).toEqual([]);
	});
});
