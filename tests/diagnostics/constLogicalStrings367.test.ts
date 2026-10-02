// Diagnostics tests: Const folding of logical operators on values outside the
// Long range, and of a string with no digit in arithmetic or a comparison
// (issue #367). Measured in Excel 16.0 (build 20326, 2026-10-01).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

function constSource(decl: string): string {
	return `Option Explicit\nPrivate Const C${decl}\nFunction Main() As Variant\n    Main = C\nEnd Function\n`;
}

function errors(src: string): string[] {
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a Const the VBE refuses', () => {
	it('overflows a logical operator outside the Long range', () => {
		for (const decl of [' = 1E10 And 1', ' = 3E9 Or 0', ' = 922337203685477@ Xor 1', ' = 3E9 Eqv 0', ' = 1E10 Imp 0', ' As Byte = False Imp 1.5', ' = 2147483648# And 1']) {
			const src = constSource(decl);
			expectDiagnostic(src, byCode(analyzeModule(src), 'const-overflow'), 'const-overflow', { message: 'Overflow' });
		}
	});

	it('is a Type mismatch for a string with no digit beside a number', () => {
		for (const decl of [' = "" + 1', ' = "abc" < 1', ' = "abc" < #1/1/2000#', ' = 1 / ""']) {
			const src = constSource(decl);
			expectDiagnostic(src, byCode(analyzeModule(src), 'const-evaluation-error'), 'const-evaluation-error', { message: 'Type mismatch' });
		}
	});

	it('compiles otherwise', () => {
		for (const decl of [' = 40000 And 1', ' = "2" * 2', ' = "a" & 1', ' = "a" = "b"', ' = 5 And 3', ' = -1 Xor 255', ' As Byte = 1 Or 2', ' = 2147483647 And 1', ' = True Or 4', ' = "a" + "b"']) {
			expect(errors(constSource(decl)), decl).toEqual([]);
		}
	});
});

describe('a logical operator at run time', () => {
	it('overflows on a value outside the Long range, and stores its result', () => {
		for (const lines of [
			['Dim n As Long', 'n = 3E9 Or 0'],
			['Dim i As Integer', 'i = 65535 Or 0'],
			// The result keeps the operands' type: Integer, and Byte for two Bytes.
			['Main = (32767 And 32767) + 1'],
			['Dim b As Byte, c As Byte', 'b = 200', 'c = 100', 'Main = (b And b) + c'],
		]) {
			const src = `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\n    Main = 1\nEnd Function\n`;
			expectDiagnostic(src, byCode(analyzeModule(src), 'arithmetic-overflow'), 'arithmetic-overflow', {});
		}
		for (const lines of [['Dim b As Byte', 'b = 5 And 3'], ['Dim i As Integer', 'i = 40000 And 1'], ['Dim i As Integer', 'i = 65535 And 255']]) {
			const src = `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\n    Main = 1\nEnd Function\n`;
			expect(byCode(analyzeModule(src), 'arithmetic-overflow'), lines.join(' : ')).toEqual([]);
		}
	});
});
