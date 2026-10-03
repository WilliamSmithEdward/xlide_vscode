// Diagnostics tests: a local's known value passed by value to the project's
// own procedure converts as a literal does (issue #558). Each case was run
// through pyVBAharness on 2026-10-02 in Excel 16.0 64-bit (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const HELPERS = 'Private Sub S(ByVal p As Byte)\nEnd Sub\n'
	+ 'Private Function F(ByVal p As Byte) As Long\n    F = p\nEnd Function\n'
	+ 'Private Sub T(ByVal a As Long, ByVal p As Byte)\nEnd Sub\n'
	+ 'Private Sub SI(ByVal p As Integer)\nEnd Sub\n'
	+ 'Private Sub SR(ByRef p As Byte)\nEnd Sub\n'
	+ 'Private Sub SO(Optional ByVal p As Byte)\nEnd Sub\n';

function errors(body: string): string[] {
	const src = `Option Explicit\n${HELPERS}Function Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

const V = 'Dim v As Integer\n    v = -3\n    ';

describe('a known value passed by value (issue #558)', () => {
	it('overflows a Byte in every call form', () => {
		for (const call of ['S v', 'S p:=v', 'Call S(v)', 'Call S(p:=v)', 'Main = F(v)', 'Main = F(p:=v)', 'S (v)', 'T 1, v', 'T p:=v, a:=1']) {
			expect(errors(`${V}${call}`), call).toEqual(['argument-type-mismatch']);
		}
	});

	it('reads a String, a Date and a parenthesized Variant', () => {
		expect(errors('Dim s As String\n    s = "abc"\n    SI s')).toEqual(['argument-type-mismatch']);
		expect(errors('Dim d As Date\n    d = #1/2/2000#\n    SI d')).toEqual(['argument-type-mismatch']);
		expect(errors('Dim v As Variant\n    v = "x"\n    SO (v)')).toEqual(['argument-type-mismatch']);
	});

	it('stays quiet where the value fits, or a ByRef passes the variable', () => {
		expect(errors('Dim v As Integer\n    v = 3\n    S v')).toEqual([]);
		expect(errors('Dim s As String\n    s = "12"\n    SI s')).toEqual([]);
		expect(errors('Dim v As Byte\n    v = 3\n    SR v')).toEqual([]);
		expect(errors('Dim v As Long\n    v = -3\n    Main = Abs(v)')).toEqual([]);
	});
});
