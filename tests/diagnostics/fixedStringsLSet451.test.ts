// Diagnostics tests: fixed-length strings and LSet/RSet (issue #451).
// Measured in Excel 16.0 64-bit (build 20430, 2026-10-02) through
// pyVBAharness.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const TYPES = 'Private Type TA\n    x As Long\nEnd Type\nPrivate Type TB\n    y As Long\nEnd Type\n';

function errors(head: string, ...lines: string[]): string[] {
	const src = `Option Explicit\n${head}Function Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => `${diag.code}: ${diag.message}`);
}

describe('LSet and RSet', () => {
	it('refuse a target they cannot fill', () => {
		for (const [decl, statement, message] of [
			['n As Long', 'LSet n = 5', 'LSet allowed only on strings and user-defined types'],
			['d As Double', 'LSet d = 5', 'LSet allowed only on strings and user-defined types'],
			['d As Date', 'LSet d = #1/1/2000#', 'LSet allowed only on strings and user-defined types'],
			['n As Long', 'RSet n = 5', 'RSet allowed only on strings'],
			['b As Boolean', 'RSet b = True', 'RSet allowed only on strings'],
		]) {
			expect(errors('', `Dim ${decl}`, statement), statement).toEqual([expect.stringMatching(new RegExp(`^lset-type-mismatch: .*${message}`))]);
		}
		expect(errors(TYPES, 'Dim a As TA, b As TB', 'RSet a = b')).toEqual([expect.stringMatching(/^lset-type-mismatch: .*RSet allowed only on strings/)]);
		expect(errors(TYPES, 'Dim a As TA', 'RSet a = "x"')).toEqual([expect.stringMatching(/^lset-type-mismatch: .*RSet allowed only on strings/)]);
	});

	it('compile on a String, a Variant, and LSet between Types of fixed size', () => {
		expect(errors('', 'Dim v As Variant', 'LSet v = "ab"')).toEqual([]);
		expect(errors('', 'Dim s As String', 's = "abcd"', 'LSet s = "x"')).toEqual([]);
		expect(errors('', 'Dim s As String * 3', 'RSet s = "x"')).toEqual([]);
		expect(errors(TYPES, 'Dim a As TA, b As TB', 'LSet a = b')).toEqual([]);
	});
});

describe('a fixed-length String', () => {
	it('needs a constant length', () => {
		expect(errors('', 'Dim L As Long', 'L = 5', 'Dim s As String * L')).toEqual([expect.stringMatching(/^array-bound-not-constant: .*Constant expression required/)]);
		expect(errors('Private Const K = 4\n', 'Dim s As String * K', 'Main = Len(s)')).toEqual([]);
	});

	it('holds Chr(0) until assigned, which converts to nothing but a String or Variant', () => {
		for (const target of ['Long', 'Double', 'Boolean', 'Date']) {
			expect(errors('', `Dim s As String * 3, x As ${target}`, 'x = s'), target).toEqual([expect.stringMatching(/^assignment-type-mismatch: .*Chr\(0\).*'13'/)]);
		}
		expect(errors('', 'Dim s As String * 3, n As Long', 's = "12"', 'n = s')).toEqual([]);
		expect(errors('', 'Dim s As String * 3, v As Variant', 'v = s')).toEqual([]);
	});
});
