// Diagnostics tests: a local declared with no type holds a Double as a
// Double, so `\`, Mod and the logical operators convert it to Long and
// overflow (issue #323). Measured in Excel 16.0 64-bit (2026-10-02)
// through pyVBAharness.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function found(lines: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n${lines}\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('an untyped local holding a Double past the Long range', () => {
	it('overflows in Mod and \\', () => {
		expect(found('Dim v\nv = 3E9\nMain = v Mod 2')).toEqual(['arithmetic-overflow']);
		expect(found('Dim v\nv = 2147483648#\nMain = v \\ 2')).toEqual(['arithmetic-overflow']);
	});

	it('stays quiet within the Long range', () => {
		expect(found('Dim v\nv = 1E9\nMain = v Mod 2')).toEqual([]);
	});

	it('is left to DefType, which may type it otherwise', () => {
		const src = 'Option Explicit\nDefLng A-Z\nFunction Main() As Variant\nDim v\nv = 5\nMain = v Mod 2\nEnd Function\n';
		expect(analyzeModule(src).filter((diag) => diag.severity === 'error')).toEqual([]);
	});
});
