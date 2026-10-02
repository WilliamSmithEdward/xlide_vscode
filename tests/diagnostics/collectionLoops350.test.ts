// Diagnostics tests: a Collection filled or emptied by a counted For loop
// (issue #350). Measured in Excel 16.0 64-bit (build 20430, 2026-10-02)
// through pyVBAharness.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(...lines: string[]): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    Dim c As New Collection, o As Collection, i As Long\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => `${diag.code}: ${diag.message}`);
}

describe('a Collection filled in a For loop', () => {
	it('raises for a key added twice, or an index past what the loop added', () => {
		expect(errors('For i = 1 To 2', 'c.Add i, "k"', 'Next')).toEqual([expect.stringMatching(/^collection-key-in-use: .*'457'/)]);
		expect(errors('For i = 1 To 3', 'c.Add i', 'Next', 'Main = c(5)')).toEqual([expect.stringMatching(/^collection-index-out-of-range: .*'9'/)]);
		expect(errors('For i = 1 To 0', 'c.Add i', 'Next', 'Main = c(1)')).toEqual([expect.stringMatching(/^collection-index-out-of-range: .*'5'/)]);
		expect(errors('c.Add 1', 'c.Add 2', 'For i = 1 To 2', 'c.Remove 1', 'Next', 'Main = c(1)')).toEqual([expect.stringMatching(/^collection-index-out-of-range: .*'5'/)]);
	});

	it('is quiet where it runs', () => {
		for (const lines of [
			['For i = 1 To 2', 'c.Add i, CStr(i)', 'Next', 'Main = c("1")'], ['For i = 1 To 3', 'c.Add i', 'Next', 'Main = c(3)'],
			['For i = 1 To 3', 'If i > 1 Then c.Add i', 'Next', 'Main = c(1)'], ['Set o = c', 'For i = 1 To 3', 'c.Add i', 'Next', 'Main = o(3)'],
		]) {
			expect(errors(...lines), lines.join(' / ')).toEqual([]);
		}
	});
});
