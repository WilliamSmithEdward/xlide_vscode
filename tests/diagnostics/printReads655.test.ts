// Diagnostics tests: Debug.Print, Debug.Assert, Print # and Write # read a
// local and leave its value known (issue #655). Measured on 2026-10-02 in
// Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const D = 'Dim a As Long, b As Long, c As Collection, f As Integer\n    ';

function errors(body: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    ${D}${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('Debug.Print and file output leave a value known (issue #655)', () => {
	it('keeps a guard on the value decided', () => {
		expect(errors('a = 1\n    Debug.Print a\n    If a = 2 Then c.Add 1')).toEqual([]);
		expect(errors('a = 1\n    If a = 1 Then\n        Debug.Print a\n    End If\n    If a = 2 Then c.Add 1')).toEqual([]);
		expect(errors('a = 1\n    If a = 1 Then Debug.Print a\n    If a = 2 Then Print #1, a')).toEqual([]);
	});

	it('keeps the value for a later division', () => {
		for (const between of ['Debug.Print a; b', 'Debug.Print a + 1', 'Debug.Print a', 'If b = 0 Then Debug.Print a', 'Print #f, a', 'Write #f, a', 'Debug.Assert a = 0']) {
			expect(errors(`a = 0\n    ${between}\n    Main = 10 \\ a`), between).toEqual(['division-by-zero']);
		}
	});

	it('still forgets a value something else writes', () => {
		expect(errors('a = 0\n    Debug.Print a\n    a = 2\n    Main = 10 \\ a')).toEqual([]);
		expect(errors('a = 0\n    Input #f, a\n    Main = 10 \\ a')).toEqual([]);
	});
});
