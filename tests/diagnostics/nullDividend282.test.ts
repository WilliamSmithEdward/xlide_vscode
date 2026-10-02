// Diagnostics tests: Null divided by zero is Null and raises nothing
// (issue #282). Measured in Excel 16.0 64-bit (build 20430, 2026-10-02)
// through pyVBAharness.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode } from '../helpers/diagnostics';

function divisions(...lines: string[]): unknown[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    Dim v As Variant, w As Variant, d As Long\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
	return byCode(analyzeModule(src), 'division-by-zero');
}

describe('a Null dividend', () => {
	it('divides by zero without raising', () => {
		for (const lines of [
			['v = Null', 'Main = IsNull(v / 0)'], ['v = Null', 'Main = IsNull(v \\ 0)'], ['v = Null', 'Main = IsNull(v Mod 0)'],
			['v = Null', 'Main = IsNull(v / d)'], ['Main = IsNull(Null / 0)'], ['v = Null', 'w = v / 0', 'Main = IsNull(w)'],
		]) {
			expect(divisions(...lines), lines.join(' / ')).toEqual([]);
		}
	});

	it('is still judged once the Variant holds a number', () => {
		expect(divisions('v = Null', 'v = 1', 'Main = v / 0')).toHaveLength(1);
		expect(divisions('Main = 1 / 0')).toHaveLength(1);
	});
});
