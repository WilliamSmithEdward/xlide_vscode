// Diagnostics tests: a String local spelling a number under ^, and a Boolean
// local in arithmetic, overflow as Excel computes them (issue #331). Each
// case was run through pyVBAharness on 2026-10-02 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(body: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a String local under ^ and a Boolean local in arithmetic (issue #331)', () => {
	it('overflow', () => {
		expect(errors('Dim s As String\n    s = "12"\n    Main = s ^ 32767')).toEqual(['arithmetic-overflow']);
		expect(errors('Dim n As Long, b As Boolean\n    n = 2147483647\n    b = True\n    Main = n - b')).toEqual(['arithmetic-overflow']);
		expect(errors('Main = "12" ^ 32767')).toEqual(['arithmetic-overflow']);
	});

	it('read a String beside True under + as a number', () => {
		expect(errors('Dim s As String\n    s = "True"\n    Main = s + True')).toEqual(['string-arithmetic-coercion']);
		expect(errors('Dim s As String\n    s = "1"\n    Main = s + True')).toEqual([]);
	});

	it('give a Date a number past its range as Type mismatch from a Variant, Overflow from a typed value (issue #329)', () => {
		const fromVariant = analyzeModule('Option Explicit\nSub T()\n    Dim v As Variant, d As Date\n    v = 2958466\n    d = v\nEnd Sub\n').filter((diag) => diag.severity === 'error');
		expect(fromVariant.map((diag) => diag.code)).toEqual(['assignment-type-mismatch']);
		expect(fromVariant[0].message).toContain("'13'");
		const typed = analyzeModule('Option Explicit\nSub T()\n    Dim d As Date\n    d = -657435\nEnd Sub\n').filter((diag) => diag.severity === 'error');
		expect(typed.map((diag) => diag.code)).toEqual(['arithmetic-overflow']);
	});

	it('stay quiet where the value fits, or + joins two Strings', () => {
		expect(errors('Dim s As String\n    s = "12"\n    Main = s ^ 2')).toEqual([]);
		expect(errors('Dim n As Long, b As Boolean\n    n = 2147483646\n    b = True\n    Main = n - b')).toEqual([]);
		expect(errors('Dim s As String, t As String\n    s = "20000"\n    t = "20000"\n    Main = s + t')).toEqual([]);
	});
});
