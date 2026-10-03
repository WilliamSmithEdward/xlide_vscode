// Diagnostics tests: a parameterless Function called with an argument list,
// which VBA reads as a call and an index into its result when that is a
// Variant or an object (issue #609). Measured on 2026-10-02 in Excel 16.0
// (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeProjectModule } from './helpers';

const HEAD = 'Private Function Arr() As Variant\n    Arr = Array(10, 20)\nEnd Function\n'
	+ 'Private Function Coll() As Collection\n    Set Coll = New Collection\n    Coll.Add 5\nEnd Function\n'
	+ 'Private Function Rng() As Range\n    Set Rng = ActiveSheet.Range("A1:B2")\nEnd Function\n'
	+ 'Private Function Untyped()\n    Untyped = Array(1, 2)\nEnd Function\n'
	+ 'Private Function LArr() As Long()\n    Dim a(1) As Long\n    LArr = a\nEnd Function\n'
	+ 'Private Function Txt() As String\n    Txt = "abc"\nEnd Function\n'
	+ 'Private Function Num() As Long\n    Num = 3\nEnd Function\n';
const MOD2 = 'Option Explicit\nPublic Function Arr2() As Variant\n    Arr2 = Array(10, 20)\nEnd Function\n';
const K = 'Option Explicit\nPublic Function Items() As Variant\n    Items = Array(10, 20)\nEnd Function\n';

function counts(body: string): string[] {
	const src = `Option Explicit\n${HEAD}Function Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeProjectModule(src, [
		{ moduleName: 'Module1', source: src },
		{ moduleName: 'Module2', source: MOD2 },
		{ moduleName: 'K', source: K, moduleKind: 'class' },
	], 'Module1').filter((d) => d.code === 'argument-count').map((d) => d.code);
}

describe('a call, then an index into its result (issue #609)', () => {
	it('takes the argument as an index on a Variant or object result', () => {
		for (const body of ['Main = Arr(1)', 'Main = Coll(1)', 'Main = Rng(1).Address', 'Main = Module2.Arr2(1)', 'Dim k As New K\n    Main = k.Items(1)',
			'Main = Untyped(1)', 'Main = Arr()(1)', 'Arr 1', 'Call Arr(1)', 'Coll 1']) {
			expect(counts(body), body).toEqual([]);
		}
	});

	it('still refuses it on a scalar or a typed array', () => {
		for (const body of ['Main = LArr(1)', 'Main = Txt(1)', 'Main = Num(1)', 'Txt 1']) {
			expect(counts(body), body).toEqual(['argument-count']);
		}
	});
});
