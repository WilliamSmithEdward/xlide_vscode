// Diagnostics tests: Dictionary and Collection leftovers (issue #349).
// Measured in Excel 16.0 64-bit (build 20430, 2026-10-02) through
// pyVBAharness.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const D = ['Dim d As Object', 'Set d = CreateObject("Scripting.Dictionary")'];
const C = ['Dim c As New Collection'];

function errors(...lines: string[]): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => `${diag.code}: ${diag.message}`);
}

describe('a Dictionary', () => {
	it('raises where the issue measured', () => {
		for (const [lines, error] of [
			[[...D, 'd.Add "a", 1', 'd.CompareMode = 1'], "'5'"],
			[[...D, 'd.CompareMode = 1', 'd.Add "k", 1', 'd.Add "K", 2'], "'457'"],
			[[...D, 'd.Key("a") = "b"'], "'32811'"],
			[[...D, 'd.Add "a", 1', 'd.Add "b", 2', 'd.Key("a") = "b"'], "'457'"],
			[[...D, 'd.Add Array(1), 1'], "'5'"],
			[[...D, 'Main = d("k").Count'], "'424'"],
		] as const) {
			expect(errors(...lines), lines.join(' / ')).toEqual([expect.stringContaining(error)]);
		}
	});

	it('is quiet where it runs', () => {
		for (const lines of [
			[...D, 'd.CompareMode = 1', 'd.Add "k", 1'], [...D, 'd.Add "a", 1', 'd.Key("a") = "b"', 'Main = d("b")'],
			[...D, 'd.Add "k", 1', 'Main = d("k")'], [...D, 'd.Add "k", Nothing', 'Set d("k") = New Collection', 'Main = d("k").Count'],
		]) {
			expect(errors(...lines), lines.join(' / ')).toEqual([]);
		}
	});
});

describe('a Collection', () => {
	it('rounds a fractional index, judges Before by key, and refuses a number held as the key', () => {
		expect(errors(...C, 'c.Add 1', 'Main = c(1.6)')).toEqual([expect.stringContaining("'9'")]);
		expect(errors(...C, 'c.Add 1, "a"', 'c.Add 2, "b", "zz"')).toEqual([expect.stringContaining("'5'")]);
		expect(errors(...C, 'Dim k As Variant', 'k = 5', 'c.Add 1, k')).toEqual([expect.stringContaining("'13'")]);
		expect(errors(...C, 'Dim k As Long', 'k = 5', 'c.Add 1, k')).toEqual([expect.stringContaining("'13'")]);
	});

	it('is quiet where it runs', () => {
		for (const lines of [
			[...C, 'c.Add 1', 'c.Add 2', 'Main = c(1.6)'], [...C, 'c.Add 1, "a"', 'c.Add 2, "b", "a"'],
			[...C, 'Dim k As String', 'k = "5"', 'c.Add 1, k', 'Main = c("5")'],
		]) {
			expect(errors(...lines), lines.join(' / ')).toEqual([]);
		}
	});
});
