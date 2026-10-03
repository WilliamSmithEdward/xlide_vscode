// Diagnostics tests: an operator on Null passed to a typed parameter (issue
// #324). Each case was run through pyVBAharness on 2026-10-02 in Excel 16.0
// (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const TAIL = 'Private Function TakeL(ByVal n As Long) As Long\n    TakeL = 1\nEnd Function\n'
	+ 'Private Function TakeS(ByVal s As String) As Long\n    TakeS = 1\nEnd Function\n'
	+ 'Private Function TakeV(ByVal v As Variant) As Long\n    TakeV = 1\nEnd Function\n'
	+ 'Private Function TakeR(n As Long) As Long\n    TakeR = 1\nEnd Function\n';

function errors(body: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n${TAIL}`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('an operator on Null passed to a typed parameter (issue #324)', () => {
	it('raises 94', () => {
		for (const body of ['Main = TakeL(1 + Null)', 'Main = TakeL(Null * 2)', 'Main = TakeS("a" + Null)', 'Main = TakeL(-Null)',
			'Main = TakeL(Null And 1)', 'Main = TakeR(1 + Null)', 'Main = TakeL(Abs(Null))', 'Dim v As Variant\n    v = Null\n    Main = TakeL(v + 1)']) {
			expect(errors(body), body).toEqual(['argument-type-mismatch']);
		}
	});

	it('stays quiet on a Variant parameter, on &, and where the other side decides', () => {
		for (const body of ['Main = TakeV(1 + Null)', 'Main = TakeS("a" & Null)', 'Main = TakeL(Null And 0)']) {
			expect(errors(body), body).toEqual([]);
		}
	});
});
