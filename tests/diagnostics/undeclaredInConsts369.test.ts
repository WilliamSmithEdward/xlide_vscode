// Diagnostics tests: a name nothing defines in a Const's value or an Enum
// member's value (issue #369). Every case was measured in Excel 16.0 (build
// 20326, 2026-10-01), compiled with the VBE's Debug > Compile.

import { describe, it, expect } from 'vitest';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';
import { analyzeProjectModule } from './helpers';

const CODE = 'undeclared-variable';

function project(src: string) {
	return analyzeProjectModule(src, [], 'Module1');
}

function source(decl: string, ...lines: string[]): string {
	return `Option Explicit\n${decl}Function Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

describe('a name nothing defines in a constant value', () => {
	it.each([
		['Private Const K = asdf\n', ['Main = K'], 'Variable not defined'],
		['', ['Const K = asdf', 'Main = K'], 'Variable not defined'],
		['Private Const K = 1 + asdf\n', ['Main = K'], 'Variable not defined'],
		['Private Enum E\n    eA\n    eB = asdf\nEnd Enum\n', ['Main = eB'], 'Constant expression required'],
	])('is reported in %j', (decl, lines, error) => {
		const src = source(decl, ...lines);
		expectDiagnostic(src, byCode(project(src), CODE), CODE, { span: 'asdf', message: error });
	});

	it('leaves names that are defined alone', () => {
		for (const decl of [
			'Private Const K = xlUp\n',
			'Private Const K = vbCrLf\n',
			'Private Const N = 2\nPrivate Const K = N * 2\n',
			'Private Const K = Excel.xlUp\n',
			'Private Const K = "asdf"\n',
			'Private Enum E\n    eA\n    eB = eA + 1\nEnd Enum\n',
			'Private Enum E\n    eA = xlUp\nEnd Enum\n',
		]) {
			expect(byCode(project(source(decl, 'Main = 1')), CODE), decl).toEqual([]);
		}
		// A module name qualifies, as issue #211 measured: Module2.M2.
		const caller = source('Private Const K = M2 + 1\nPrivate Const J = Module2.M2 + 1\n', 'Main = K + J');
		const diagnostics = analyzeProjectModule(caller, [{ moduleName: 'Module2', source: 'Option Explicit\nPublic Const M2 = 3\n' }], 'Module1');
		expect(byCode(diagnostics, CODE)).toEqual([]);
	});

	it('needs Option Explicit', () => {
		const src = 'Private Const K = asdf\nFunction Main() As Variant\n    Main = K\nEnd Function\n';
		expect(byCode(project(src), CODE)).toEqual([]);
	});
});
