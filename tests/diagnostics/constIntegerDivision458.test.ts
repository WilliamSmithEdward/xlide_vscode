// Diagnostics tests: `\` and Mod convert their operands to a Long, and a
// string operand of a logical operator, `\` or Mod is read as the number it
// spells (issue #458). Measured in Excel 16.0 64-bit (build 20430,
// 2026-10-02) through pyVBAharness.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

function constSource(decl: string): string {
	return `Option Explicit\nPrivate Const C${decl}\nFunction Main() As Variant\n    Main = C\nEnd Function\n`;
}

function runSource(line: string): string {
	return `Option Explicit\nFunction Main() As Variant\n    ${line}\nEnd Function\n`;
}

function errors(src: string): string[] {
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a Const the VBE refuses', () => {
	it('overflows \\ or Mod with an operand outside the Long range', () => {
		for (const decl of [
			' = 65535 \\ 3.4E+38', ' = 1E10 \\ 2', ' = 7 Mod 1E10', ' = 3E10 Mod 7', ' = 1 \\ 922337203685477@',
			' = 255 Mod 922337203685477@', ' = 2147483648# Mod 3', ' = 100 \\ "1E10"',
			' = "3E9" Or 0', ' As Byte = &O777 Imp "1E3"', ' As Integer = "40000" Or 0',
		]) {
			const src = constSource(decl);
			expectDiagnostic(src, byCode(analyzeModule(src), 'const-overflow'), 'const-overflow', { message: 'Overflow' });
		}
	});

	it('divides by zero with a string dividend', () => {
		const src = constSource(' = "1" \\ False');
		expectDiagnostic(src, byCode(analyzeModule(src), 'const-evaluation-error'), 'const-evaluation-error', { message: 'Division by zero' });
	});

	it('compiles otherwise', () => {
		for (const decl of [' = 7 \\ 2', ' = 7 Mod 2', ' = 2147483647 \\ 1', ' = "12" \\ 5', ' = "12" Mod 5', ' = "3" Or 4', ' = 2147483647.4 \\ 1', ' = 6 \\ "2" * 3']) {
			expect(errors(constSource(decl)), decl).toEqual([]);
		}
	});
});

describe('the same at run time', () => {
	it('overflows, or divides by zero by False', () => {
		for (const line of ['Main = 1E10 \\ 2', 'Main = 3E10 Mod 7', 'Main = "3E9" Or 0']) {
			const src = runSource(line);
			expectDiagnostic(src, byCode(analyzeModule(src), 'arithmetic-overflow'), 'arithmetic-overflow', { message: "'6'" });
		}
		for (const line of ['Main = 1 \\ False', 'Main = 2.5 / False', 'Main = 5 Mod False', 'Main = "1" \\ False']) {
			const src = runSource(line);
			expectDiagnostic(src, byCode(analyzeModule(src), 'division-by-zero'), 'division-by-zero', { message: "'11'" });
		}
		expect(errors(runSource('Main = "12" \\ 5'))).toEqual([]);
	});
});
