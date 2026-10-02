// Diagnostics tests: Int and Fix of a LongLong under Mod, and a Variant
// holding a Double (issue #480). Each raising sample was measured through
// pyVBAharness on 2026-10-02 in Excel 16.0 (build 20430); each quiet one
// runs there.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode } from '../helpers/diagnostics';

const OVERFLOW = 'arithmetic-overflow';

function wrap(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

describe('Mod and \\ on a LongLong or a Variant (issue #480)', () => {
	it('keeps a LongLong through Int and Fix', () => {
		for (const use of ['Main = CStr(Fix(q) Mod 7)', 'Main = CStr(Int(q) Mod 7)', 'Main = CStr(Fix(q) \\ 2)', 'Main = CStr(q Mod 7)']) {
			expect(byCode(analyzeModule(wrap('Dim q As LongLong', 'q = 2147483648^', use)), OVERFLOW), use).toHaveLength(0);
		}
	});

	it('reports a Variant holding a Double past the Long range', () => {
		for (const use of ['Main = CStr(v Mod 7)', 'Main = v \\ 2']) {
			expect(byCode(analyzeModule(wrap('Dim v As Variant', 'v = 2147483648#', use)), OVERFLOW), use).toHaveLength(1);
		}
		expect(byCode(analyzeModule(wrap('Dim d As Double', 'd = 2147483648#', 'Main = CStr(Fix(d) Mod 7)')), OVERFLOW)).toHaveLength(1);
	});

	it('widens a Variant\'s arithmetic instead of overflowing', () => {
		const bodies = [
			['Dim v As Variant', 'v = 32767', 'Main = v + 1'],
			['Dim v', 'v = 2147483647', 'Main = v + 1'],
			['Dim v As Variant', 'v = 32767', 'v = v * 2', 'Main = v'],
			['Dim v As Variant', 'v = 2147483648^', 'Main = CStr(v Mod 7)'],
		];
		for (const body of bodies) {
			expect(byCode(analyzeModule(wrap(...body)), OVERFLOW), body.join(' / ')).toHaveLength(0);
		}
	});

	it('still converts a Variant\'s value where a conversion or a typed target needs it', () => {
		const bodies = [
			['Dim v As Variant', 'v = 40000', 'Main = CInt(v)'],
			['Dim v As Variant', 'v = 255', 'Main = CByte(v + 1)'],
			['Dim v As Variant, i As Integer', 'v = 40000', 'i = v'],
		];
		for (const body of bodies) {
			expect(byCode(analyzeModule(wrap(...body)), OVERFLOW), body.join(' / ')).toHaveLength(1);
		}
	});
});
