// Diagnostics tests: the column letter Split takes from a literal Range's
// Address (issue #457). Each case was run through pyVBAharness on 2026-10-02
// in Excel 16.0 64-bit (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(type: string, index: number): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    Dim x As ${type}\n    x = Split(Range("C1").Address, "$")(${index})\n    Main = x\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('Split of a literal Range\'s Address (issue #457)', () => {
	it('gives the column letter, which no number type takes', () => {
		for (const type of ['Long', 'Double', 'Boolean', 'Date']) {
			expect(errors(type, 1)).toHaveLength(1);
		}
	});

	it('gives the row, which converts, and stays quiet into a String', () => {
		for (const type of ['Long', 'Double', 'Boolean', 'Date']) {
			expect(errors(type, 2)).toEqual([]);
		}
		expect(errors('String', 1)).toEqual([]);
	});
});
