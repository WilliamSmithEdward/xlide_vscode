// Diagnostics tests: Split of a string expression, and starts folded from
// InStr and Len of one (issue #509). Each sample was measured through
// pyVBAharness on 2026-10-02 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(expr: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    Main = ${expr}\nEnd Function\n`;
	return analyzeModule(src).filter((d) => d.severity === 'error').map((d) => /Run-time error '(\d+)'/.exec(d.message)?.[1] ?? d.code ?? '');
}

describe('Split of a string expression (issue #509)', () => {
	it('indexes the parts of a folded string', () => {
		for (const expr of [
			'Split(Mid("12", 5), ",")(0)',
			'Split(LCase("a,b"), ",")(2)',
			'Split("ab" & "c", ",")(2)',
			'Split(StrReverse("12"), ",")(2)',
			'Split(Left("a,b,c", 3), ",")(2)',
			'Split(Replace("a;b", ";", ","), ",")(2)',
			'Split(Space(2) & "x", " ")(3)',
		]) {
			expect(errors(expr), expr).toEqual(['9']);
		}
	});

	it('folds InStr and Len of an expression into a start', () => {
		expect(errors('Mid("12", InStr("", ""))')).toEqual(['5']);
		expect(errors('Mid("", Len(UCase("")))')).toEqual(['5']);
	});

	it('stays quiet where the index or start holds', () => {
		for (const expr of [
			'Split(LCase("a,b"), ",")(1)',
			'Split("ab" & ",c", ",")(1)',
			'Split(Right("x,y,z", 3), ",")(1)',
			'Mid("12", InStr("ab", "b"))',
			'Split(Trim("  a,b  "), ",")(1)',
			'InStr("", "")',
			'Split(UCase$("p,q"), ",")(1)',
		]) {
			expect(errors(expr), expr).toEqual([]);
		}
	});
});
