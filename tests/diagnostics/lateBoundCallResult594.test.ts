// Diagnostics tests: a late-bound call's result is not its argument (issue
// #594). Measured on 2026-10-02 in Excel 16.0 (build 20430): each runs.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function memberNotFound(body: string): string[] {
	const source = `Option Explicit\nFunction Main() As Variant\n    ${body}\nEnd Function\n`;
	return analyzeModule(source).filter((d) => d.code === 'runtime-member-not-found').map((d) => d.code);
}

describe('a late-bound call that takes a Range (issue #594)', () => {
	it('is not taken for the Range', () => {
		for (const body of [
			'Dim k As Object\n    Main = k.Wrap(ActiveSheet.Range("A1")).Caption',
			'Dim k As Variant\n    Main = k.Wrap(ActiveSheet.Range("A1")).Caption',
			'Dim k As Object\n    Main = k.Wrap(ActiveSheet.Range("A1")) _\n        .Caption',
			'Dim k As Object\n    Main = k.Wrap(Range("A1")).Caption',
			'Dim k As Object\n    Main = k.Wrap(Cells(1, 1)).Caption',
		]) {
			expect(memberNotFound(body), body).toEqual([]);
		}
	});

	it('still reads a Range in grouping parentheses', () => {
		expect(memberNotFound('Main = (Range("A1")).Nope')).toEqual(['runtime-member-not-found']);
		expect(memberNotFound('Dim r As Range\n    Set r = Range("A1")\n    Main = r.Offset(1).Nope')).toEqual(['runtime-member-not-found']);
	});
});
