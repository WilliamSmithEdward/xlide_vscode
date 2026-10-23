// Diagnostics tests: a Function that needs an argument read as a value with
// none, and VBA functions a parameterless project Function of the same name
// does not hide (issue #645). Measured on 2026-10-02 in Excel 16.0 (build
// 20430).

import { describe, it, expect } from 'vitest';
import { analyzeProjectModule } from './helpers';

const F = 'Private Function F(ByVal x As Variant) As Long\n    F = 7\nEnd Function\n';
const LEFT0 = 'Private Function Left() As Long\n    Left = 7\nEnd Function\n';

function errors(decls: string, body: string, module2?: string): string[] {
	const src = `Option Explicit\n${decls}Function Main() As Variant\n    ${body}\nEnd Function\n`;
	const modules = [{ moduleName: 'Module1', source: src }, ...(module2 ? [{ moduleName: 'Module2', source: `Option Explicit\n${module2}` }] : [])];
	return analyzeProjectModule(src, modules, 'Module1').filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a Function needing an argument, read with none (issue #645)', () => {
	it('reports it read bare', () => {
		for (const body of ['Main = F', 'Main = F + 1', 'Main = Module1.F', 'Main = CStr(F)']) {
			expect(errors(F, body), body).toEqual(['argument-count']);
		}
		expect(errors('Private Function Name(ByVal x As Variant, Optional ByVal y As Variant) As Long\n    Name = 3\nEnd Function\n', 'Main = Name')).toEqual(['argument-count']);
		expect(errors('Private Property Get P(ByVal i As Long) As Long\n    P = i\nEnd Property\n', 'Main = P')).toEqual(['argument-count']);
		const h = 'Public Function H(ByVal x As Variant) As Long\n    H = 7\nEnd Function\n';
		expect(errors('', 'Main = H', h)).toEqual(['argument-count']);
		expect(errors('', 'Main = Module2.H', h)).toEqual(['argument-count']);
	});

	it('stays quiet with an argument, an Optional parameter, a local or the Function itself', () => {
		expect(errors(F, 'Main = F(1)')).toEqual([]);
		expect(errors('Private Function G(Optional ByVal x As Variant) As Long\n    G = 8\nEnd Function\n', 'Main = G + 1')).toEqual([]);
		expect(errors(F, 'Dim F As Long\n    F = 2\n    Main = F')).toEqual([]);
		// AddressOf takes the procedure's address, not its value.
		const take = 'Private Function Take(ByVal p As LongPtr) As LongPtr\n    Take = p\nEnd Function\n';
		expect(errors(F + take, 'Main = Take(AddressOf F)').filter((code) => code === 'argument-count')).toEqual([]);
		expect(errors(F + take, 'Main = Take(AddressOf Module1.F)').filter((code) => code === 'argument-count')).toEqual([]);
		expect(errors('Private Function K(ByVal x As Long) As Long\n    K = x\n    If x > 0 Then K = K + 1\nEnd Function\n', 'Main = K(2)')).toEqual([]);
	});

	it('lets Left, InStr and StrComp past a parameterless project Function', () => {
		for (const body of ['Main = Left("abc", 1)', 'Main = Left', 'Main = Left$("abc", 1)', 'Main = InStr("abc", "c")']) {
			expect(errors(LEFT0, body), body).toEqual([]);
		}
		expect(errors(LEFT0.replace(/Left/g, 'InStr'), 'Main = InStr(1, "abc", "c")')).toEqual([]);
		expect(errors(LEFT0.replace(/Left/g, 'StrComp'), 'Main = StrComp("a", "b")')).toEqual([]);
		// One or three arguments, or another name, is refused.
		expect(errors(LEFT0, 'Main = Left("abc")')).toEqual(['argument-count']);
		expect(errors(LEFT0.replace(/Left/g, 'Right'), 'Main = Right("abc", 1)')).toEqual(['argument-count']);
		expect(errors('Private Function Val() As Long\n    Val = 1\nEnd Function\n', 'Main = Val("12")')).toEqual(['argument-count']);
	});
});
