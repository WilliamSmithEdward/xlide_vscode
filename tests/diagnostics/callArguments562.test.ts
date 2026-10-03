// Diagnostics tests: a Function of the module run for a call's literal
// arguments, and one that returns a module variable nothing writes (issue
// #562). Each case was run through pyVBAharness on 2026-10-02 in Excel 16.0
// (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const HELPERS = 'Private m As Long\n'
	+ 'Private Function Sign1(ByVal n As Long) As Long\n    If n > 0 Then Sign1 = 1 Else Sign1 = 0\nEnd Function\n'
	+ 'Private Function OptDef(Optional ByVal n As Long = 4) As Long\n    OptDef = n\nEnd Function\n'
	+ 'Private Function MaybeColl(ByVal b As Boolean) As Collection\n    If b Then Set MaybeColl = New Collection\nEnd Function\n'
	+ 'Private Function EarlyFive(ByVal b As Boolean) As Long\n    If b Then\n        EarlyFive = 5\n        Exit Function\n    End If\nEnd Function\n'
	+ 'Private Function Twice(ByVal n As Long) As Long\n    Twice = n * 2\nEnd Function\n'
	+ 'Private Function Pick(ByVal n As Long) As Long\n    Select Case n\n    Case 1\n        Pick = 0\n    Case Else\n        Pick = 3\n    End Select\nEnd Function\n'
	+ 'Private Function GetMod() As Long\n    GetMod = m\nEnd Function\n';

function errors(body: string): string[] {
	const src = `Option Explicit\n${HELPERS}Function Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a Function run for its arguments (issue #562)', () => {
	it('reports a result of 0 the arguments give', () => {
		for (const call of ['Sign1(-1)', 'OptDef(0)', 'EarlyFive(False)', 'Twice(0)', 'Pick(1)']) {
			expect(errors(`Main = 10 / ${call}`)).toEqual(['division-by-zero']);
		}
		expect(errors('Dim n As Long\n    n = 10 \\ Sign1(-5)')).toEqual(['division-by-zero']);
		expect(errors('Main = 10 Mod OptDef(0)')).toEqual(['division-by-zero']);
	});

	it('stays quiet where the arguments give another result', () => {
		for (const call of ['Sign1(3)', 'OptDef()', 'EarlyFive(True)', 'Pick(2)']) {
			expect(errors(`Main = 10 / ${call}`)).toEqual([]);
		}
	});

	it('reports an object result the arguments leave Nothing', () => {
		expect(errors('Main = MaybeColl(False).Count')).toEqual(['object-variable-not-set']);
		expect(errors('Main = MaybeColl(True).Count')).toEqual([]);
	});

	it('reads a module variable nothing writes', () => {
		expect(errors('Main = 10 / GetMod()')).toEqual(['division-by-zero']);
		expect(errors('m = 5\n    Main = 10 / GetMod()')).toEqual([]);
	});
});
