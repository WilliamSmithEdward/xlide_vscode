// Diagnostics tests: AddressOf where the VBE refuses it (issue #299).
// Measured in 64-bit Excel 16.0 (2026-10-02) through pyVBAharness.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const HELPERS = 'Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long\n'
	+ 'Public Function Cb(ByVal a As LongPtr) As LongPtr\n    Cb = a\nEnd Function\n'
	+ 'Private Function Take(ByVal p As LongPtr) As LongPtr\n    Take = p\nEnd Function\n'
	+ 'Private Function TakeL(ByVal p As Long) As Long\n    TakeL = p\nEnd Function\n'
	+ 'Private Sub PrivSub()\nEnd Sub\n';

function found(body: string): string[] {
	const src = `Option Explicit\n${HELPERS}Function Main() As Variant\n    ${body}\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => `${diag.code}: ${diag.message}`);
}

describe('AddressOf', () => {
	it('stands only as a whole argument of a project procedure or a method', () => {
		expect(found('Dim p As LongPtr\n    p = AddressOf Cb')).toEqual([expect.stringMatching(/^addressof-misuse: .*Syntax error/)]);
		expect(found('Main = Take((AddressOf Cb))')).toEqual([expect.stringMatching(/Syntax error/)]);
		expect(found('Main = Take(AddressOf Cb + 1)')).toEqual([expect.stringMatching(/Argument not optional/)]);
		expect(found('Main = CStr(AddressOf Cb)')).toEqual([expect.stringMatching(/Syntax error/)]);
	});

	it('is refused by some VBA functions and by Debug.Print', () => {
		for (const fn of ['Len', 'CLng', 'CLngPtr', 'Abs']) {
			expect(found(`Main = ${fn}(AddressOf Cb)`), fn).toEqual([expect.stringMatching(/Syntax error/)]);
		}
		expect(found('Main = ObjPtr(AddressOf Cb)')).toEqual([expect.stringMatching(/Type mismatch/)]);
		expect(found('Debug.Print AddressOf Cb')).toEqual([expect.stringMatching(/Syntax error/)]);
		expect(found('Dim x As Long\n    Debug.Print AddressOf x')).toEqual([expect.stringMatching(/Syntax error/)]);
		expect(found('Debug.Print ObjPtr(AddressOf Cb)')).toEqual([expect.stringMatching(/Type mismatch/)]);
	});

	it('is taken by the other VBA functions', () => {
		for (const fn of ['VarPtr', 'StrPtr', 'Hex', 'IsEmpty', 'TypeName']) {
			expect(found(`Main = ${fn}(AddressOf Cb)`), fn).toEqual([]);
		}
	});

	it('takes only a procedure of the project', () => {
		expect(found('Main = Take(AddressOf NoSuchProc)')).toEqual([expect.stringMatching(/Variable not defined/)]);
		expect(found('Dim v As Long\n    Main = Take(AddressOf v)')).toEqual([expect.stringMatching(/Expected Sub, Function, or Property/)]);
		expect(found('Main = Take(AddressOf Len)')).toEqual([expect.stringMatching(/^reserved-keyword-in-expression: .*Syntax error/)]);
		expect(found('Main = Take(AddressOf Sqr)')).toEqual([expect.stringMatching(/^addressof-misuse: .*Invalid use of AddressOf operator/)]);
		expect(found('Main = Take(AddressOf Mid)')).toEqual([expect.stringMatching(/^addressof-misuse: .*Invalid use of AddressOf operator/)]);
		expect(found('Main = Take(AddressOf GetTickCount)')).toEqual([expect.stringMatching(/Invalid use of AddressOf operator/)]);
	});

	it('gives a LongPtr, which a ByVal Long takes no more than a LongLong', () => {
		expect(found('Main = TakeL(AddressOf Cb)')).toEqual([expect.stringMatching(/Type mismatch/)]);
	});

	it('compiles where it belongs', () => {
		for (const body of ['Main = Take(AddressOf Cb)', 'Main = Take(AddressOf PrivSub)', 'Dim c As New Collection\n    c.Add AddressOf Cb\n    Main = c.Count', 'Take AddressOf Cb']) {
			expect(found(body), body).toEqual([]);
		}
	});
});
