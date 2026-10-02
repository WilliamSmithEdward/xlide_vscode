// Diagnostics tests: what a Collection holds, followed into its items: an
// inner Collection it shares, a number or a string (issue #452). Measured in
// Excel 16.0 64-bit (build 20430, 2026-10-02) through pyVBAharness.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const SETUP = ['Dim c As New Collection, inner As New Collection', 'inner.Add 10', 'c.Add inner, "in"', 'c.Add 5, "n"'];

function errors(...lines: string[]): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n${[...SETUP, ...lines].map((line) => `    ${line}`).join('\n')}\nEnd Function\nPrivate Sub Fill(x As Collection)\n    x.Add 99\nEnd Sub\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => `${diag.code}: ${diag.message}`);
}

describe('an item that is a Collection', () => {
	it('is judged by what the inner Collection holds', () => {
		for (const lines of [['Main = c(1)(2)'], ['Main = c("in")(3)'], ['c(1).Remove 2'], ['Main = c.Item(1).Item(2)'], ['Main = c(1).Item(2)']]) {
			expect(errors(...lines), lines.join(' / ')).toEqual([expect.stringMatching(/^collection-index-out-of-range: .*'9'/)]);
		}
		expect(errors('c(1).Remove 1', 'Main = c(1)(1)')).toEqual([expect.stringMatching(/^collection-index-out-of-range: .*'5'/)]);
		expect(errors('Main = c("in")("x")')).toEqual([expect.stringMatching(/^collection-key-not-found: /)]);
		expect(errors('c(1).Add 1, "k"', 'c(1).Add 2, "k"')).toEqual([expect.stringMatching(/^collection-key-in-use: .*'457'/)]);
	});

	it('raises 450 Let into a value', () => {
		for (const lines of [['Dim v As Long', 'v = c(1)'], ['Dim v As Variant', 'v = c(1)'], ['Main = c(1)']]) {
			expect(errors(...lines), lines.join(' / ')).toEqual([expect.stringMatching(/^object-default-value: .*'450'/)]);
		}
	});

	it('follows changes made through either name', () => {
		for (const lines of [
			['Main = c(1).Count'], ['Main = c(1)(1)'], ['c(1).Add 20', 'Main = c(1)(2)'], ['c(1).Add 20', 'Main = inner(2)'],
			['inner.Add 20', 'Main = c(1)(2)'], ['Set inner = New Collection', 'Main = c(1)(1)'], ['c(1).Remove 1', 'Main = c(1).Count'],
			['Dim o As Collection', 'Set o = c(1)', 'Main = o.Count'], ['Fill inner', 'Main = c(1)(2)'],
			['If c.Count > 0 Then inner.Add 30', 'Main = c(1)(2)'],
		]) {
			expect(errors(...lines), lines.join(' / ')).toEqual([]);
		}
	});
});

describe('an item that is a number or a string', () => {
	it('raises 424 for a member and 13 for an index', () => {
		for (const lines of [['Main = c(2).Count'], ['c(2).Add 1'], ['Main = c("n").Count'], ['c.Add "s", "s"', 'Main = c(3).Count']]) {
			expect(errors(...lines), lines.join(' / ')).toEqual([expect.stringMatching(/^variant-value-misuse: .*'424'/)]);
		}
		for (const lines of [['Main = c(2)(1)'], ['c.Add "s", "s"', 'Main = c(3)(1)']]) {
			expect(errors(...lines), lines.join(' / ')).toEqual([expect.stringMatching(/^variant-value-misuse: .*'13'/)]);
		}
	});

	it('is a value otherwise', () => {
		for (const lines of [['Main = c(2) + 1'], ['Main = c(1)(1) + c(2)'], ['Dim s As String', 's = c(2)']]) {
			expect(errors(...lines), lines.join(' / ')).toEqual([]);
		}
	});
});
