// Diagnostics tests: what the state rules know is carried into blocks (issue
// #237). Measured through pyVBAharness in Excel 16.0 on 2026-09-30: each
// error raises inside each kind of block below, which never touches the
// variable. tests/oracleWrappedInIf.test.ts checks the same over the oracle.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode } from '../helpers/diagnostics';

const ERRORS: ReadonlyArray<[string, string[], string, string]> = [
	['object-variable-not-set', ['Dim o As Collection'], 'Main = o.Count', 'never set'],
	['object-variable-not-set', ['Dim o As Collection', 'Set o = Nothing'], 'Main = o.Count', 'set to Nothing'],
	['object-variable-not-set', ['Dim o As Collection', 'Set o = Nothing'], 'With o\nMain = .Count\nEnd With', 'Nothing under With'],
	['collection-key-not-found', ['Dim c As New Collection', 'c.Add 1, "k1"'], 'Main = c("nokey")', 'a missing key'],
	['collection-index-out-of-range', ['Dim c As New Collection', 'c.Add 1'], 'Main = c(5)', 'an index past the end'],
	['collection-key-in-use', ['Dim c As New Collection', 'c.Add 1, "k1"'], 'c.Add 2, "k1"', 'a key in use'],
	['unallocated-dynamic-array-access', ['Dim dyn2() As Long', 'ReDim dyn2(2)', 'Erase dyn2'], 'Main = dyn2(0)', 'an erased array'],
	['unallocated-dynamic-array-access', ['Dim dyn2() As Long'], 'Main = dyn2(0)', 'an array never allocated'],
];

const BLOCKS: ReadonlyArray<[string, string, string]> = [
	['If True Then', 'If True Then', 'End If'],
	['If n3 = 0 Then', 'If n3 = 0 Then', 'End If'],
	['For', 'For zi = 1 To 1', 'Next'],
	['Select Case', 'Select Case 1\nCase 1', 'End Select'],
	['Do', 'Do', 'Exit Do\nLoop'],
	['With', 'With New Collection', 'End With'],
];

function main(lines: readonly string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.join('\n')}\nEnd Function\n`;
}

describe('state rules inside blocks (issue #237)', () => {
	for (const [code, setup, line, what] of ERRORS) {
		it.each(BLOCKS)(`${code}: ${what} inside %s`, (_label, open, close) => {
			const src = main(['Dim n3 As Long, zi As Long', ...setup, open, line, close]);
			expect(byCode(analyzeModule(src), code)).toHaveLength(1);
		});
	}
});

describe('what a block changes is not carried into it (issue #237)', () => {
	it.each([
		['a key a With adds', 'collection-key-not-found', ['Dim outer As New Collection', 'Dim inner As New Collection', 'With outer', '.Add Item:="x", Key:="x"', 'With inner', '.Add Item:=outer.Item("x"), Key:="copied"', 'End With', 'End With']],
		['an object an If condition sets ByRef', 'object-variable-not-set', ['Dim o As Collection', 'If TryGet(o) Then', 'Main = o.Count', 'End If']],
		['a loop that sets the object on a later pass', 'object-variable-not-set', ['Dim o As Collection, i As Long', 'For i = 1 To 2', 'If i = 2 Then Main = o.Count', 'Set o = New Collection', 'Next']],
		['a For Each control variable', 'object-variable-not-set', ['Dim o As Object, c As New Collection', 'c.Add New Collection', 'For Each o In c', 'Main = o.Count', 'Next']],
		['a key another Case adds', 'collection-key-in-use', ['Dim c As New Collection', 'Select Case Main', 'Case 1', 'c.Add 1, "k"', 'Case 2', 'c.Add 2, "k"', 'End Select']],
	])('leaves %s alone', (_label, code, lines) => {
		const src = `${main(lines)}Private Function TryGet(ByRef o As Collection) As Boolean\n    Set o = New Collection\n    TryGet = True\nEnd Function\n`;
		expect(byCode(analyzeModule(src), code)).toEqual([]);
	});

	it.each([
		['a block If', ['If i = 2 Then', 'c.Add 3, "k"', 'End If']],
		['a single-line If', ['If i = 2 Then c.Add 3, "k"']],
	])('leaves a key %s in a loop adds once an earlier pass removed it', (_label, add) => {
		const src = main(['Dim c As New Collection, i As Long', 'c.Add 1, "k"', 'For i = 1 To 2', ...add, 'If i = 1 Then c.Remove "k"', 'Next']);
		expect(byCode(analyzeModule(src), 'collection-key-in-use')).toEqual([]);
	});

	it('reports an overflow on a loop\'s first pass', () => {
		const src = main(['Dim n As Integer, i As Long', 'n = 32767', 'For i = 1 To 2', 'n = n + 1', 'Next']);
		expect(byCode(analyzeModule(src), 'arithmetic-overflow')).toHaveLength(1);
	});

	it('reports a key added twice in one pass of a loop', () => {
		const src = main(['Dim c As New Collection, i As Long', 'For i = 1 To 2', 'c.Add 1, "k"', 'c.Add 2, "k"', 'Next']);
		expect(byCode(analyzeModule(src), 'collection-key-in-use')).toHaveLength(1);
	});
});
