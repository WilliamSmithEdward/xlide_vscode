// Diagnostics tests: arguments to project procedures (issue #410).
// Measured in Excel 16.0 64-bit (build 20430, 2026-10-02) with Debug >
// Compile.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const PROCS = 'Private Sub TL(p As Long)\nEnd Sub\nPrivate Sub TS(p As String)\nEnd Sub\nPrivate Sub TC(p As Collection)\nEnd Sub\n'
	+ 'Private Sub TArr(p() As Long)\nEnd Sub\nPrivate Sub TTwo(a As Long, b As Long)\nEnd Sub\n'
	+ 'Private Sub TOptTwo(a As Long, Optional b As Long)\nEnd Sub\nPrivate Function Add2(a As Long, b As Long) As Long\n    Add2 = a + b\nEnd Function\n';

function errors(call: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    Dim c As New Collection, n As Long, s As String, v As Variant, arr() As Long\n    ${call}\nEnd Function\n${PROCS}`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => `${diag.code}: ${diag.message}`);
}

describe('arguments the VBE refuses', () => {
	it('a Collection and a scalar swapped ByRef', () => {
		for (const call of ['TL c', 'TS c', 'TC n', 'TC s', 'TC v']) {
			expect(errors(call), call).toEqual([expect.stringMatching(/^byref-argument-type-mismatch: .*ByRef argument type mismatch/)]);
		}
	});

	it('a literal or a Collection for an array parameter', () => {
		for (const call of ['TArr -1', 'TArr "a"', 'TArr c']) {
			expect(errors(call), call).toEqual([expect.stringMatching(/^argument-shape-mismatch: .*array or user-defined type expected/)]);
		}
	});

	it('a required argument named nowhere', () => {
		for (const call of ['TTwo a:=1', 'TTwo b:=2', 'Main = Add2(b:=2)']) {
			expect(errors(call), call).toEqual([expect.stringMatching(/^argument-count: .*Argument not optional/)]);
		}
	});
});

describe('arguments that compile', () => {
	it('are quiet', () => {
		for (const call of ['TTwo 1, b:=2', 'TTwo b:=2, a:=1', 'TOptTwo a:=1', 'TC c', 'TL n', 'TArr arr', 'TC Nothing']) {
			expect(errors(call), call).toEqual([]);
		}
	});
});
