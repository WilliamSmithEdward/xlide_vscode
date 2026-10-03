// Diagnostics tests: a value where a Collection parameter takes an object
// (issue #410). Each case was run through pyVBAharness on 2026-10-02 in Excel
// 16.0 64-bit (build 20430), positional and named.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(passing: 'ByRef' | 'ByVal', body: string): string[] {
	const src = `Option Explicit\nPrivate Function Callee(${passing} p As Collection) As Long\n    Callee = 1\nEnd Function\nFunction Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a value into a Collection parameter (issue #410)', () => {
	it('does not compile as True, a date, a negative number or a scalar expression', () => {
		for (const passing of ['ByRef', 'ByVal'] as const) {
			for (const arg of ['True', '#1/2/2000#', '-5']) {
				expect(errors(passing, `Main = Callee(${arg})`)).toEqual(['argument-object-type-mismatch']);
				expect(errors(passing, `Main = Callee(p:=${arg})`)).toEqual(['argument-object-type-mismatch']);
			}
			expect(errors(passing, 'Dim v As Long\n    Main = Callee(v + 0)')).toEqual(['argument-object-type-mismatch']);
		}
	});

	it('does not compile as a scalar variable passed ByVal', () => {
		for (const type of ['Long', 'String', 'Currency']) {
			expect(errors('ByVal', `Dim v As ${type}\n    Main = Callee(v)`)).toEqual(['argument-object-type-mismatch']);
			expect(errors('ByVal', `Dim v As ${type}\n    Main = Callee(p:=v)`)).toEqual(['argument-object-type-mismatch']);
		}
	});

	it('raises 424 for a Variant holding no object, passed ByVal', () => {
		expect(errors('ByVal', 'Dim v As Variant\n    Main = Callee(v)')).toEqual(['argument-type-mismatch']);
		expect(errors('ByVal', 'Dim v As Variant\n    v = 5\n    Main = Callee(p:=v)')).toEqual(['argument-type-mismatch']);
	});

	it('reads a Variant parameter or a Static local as the object it may hold', () => {
		// The shape of ModernJsonInVBA's Json_Serializer, which the corpus ran.
		const src = 'Option Explicit\nPrivate Sub Take(ByVal obj As Collection)\nEnd Sub\nPrivate Sub Pass(ByRef v As Variant)\n    Take v\nEnd Sub\nPrivate Sub Kept()\n    Static v As Variant\n    Take v\nEnd Sub\n';
		expect(analyzeModule(src).filter((diag) => diag.severity === 'error')).toEqual([]);
	});

	it('takes a Collection', () => {
		expect(errors('ByVal', 'Dim v As New Collection\n    Main = Callee(v)')).toEqual([]);
		expect(errors('ByRef', 'Dim v As New Collection\n    Main = Callee(p:=v)')).toEqual([]);
		expect(errors('ByVal', 'Dim v As Variant, c As New Collection\n    For Each v In c\n        Main = Callee(v)\n    Next')).toEqual([]);
	});
});
