// Diagnostics tests: a Type array field sized by a Const or an Enum member is
// a fixed array (issue #366). Every case was measured in Excel 16.0 (build
// 20326, 2026-10-01).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';
import { analyzeProjectModule } from './helpers';

const CONST_TYPE = 'Private Const N = 3\nPrivate Type Row3\n    vals(1 To N) As Long\nEnd Type\n';

function source(decl: string, ...lines: string[]): string {
	return `Option Explicit\n${decl}\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

function errors(src: string): string[] {
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a Type array field sized by a Const or Enum member', () => {
	it('is not read as empty', () => {
		for (const [decl, lines] of [
			[CONST_TYPE, ['Dim r As Row3', 'Main = UBound(r.vals)']],
			[CONST_TYPE, ['Dim r As Row3', 'Main = LBound(r.vals)']],
			[CONST_TYPE, ['Dim r As Row3, i As Long', 'For i = 1 To N', '    r.vals(i) = i', 'Next', 'Main = r.vals(3)']],
			[CONST_TYPE, ['Dim r As Row3', 'With r', '    Main = UBound(.vals)', 'End With']],
			['Private Const N = 3\nPrivate Type Row3\n    vals(N) As Long\nEnd Type\n', ['Dim r As Row3', 'r.vals(3) = 5', 'Main = r.vals(3)']],
			['Private Const N = 3\nPrivate Type Row3\n    vals(0 To N - 1) As Long\nEnd Type\n', ['Dim r As Row3', 'r.vals(2) = 5', 'Main = r.vals(2)']],
			['Private Enum Sz\n    szMax = 3\nEnd Enum\nPrivate Type Row3\n    vals(1 To szMax) As Long\nEnd Type\n', ['Dim r As Row3', 'Main = UBound(r.vals)']],
			['Private Const N = 2\nPrivate Type Row3\n    vals(N * 2) As Long\nEnd Type\n', ['Dim r As Row3', 'Main = UBound(r.vals)']],
			['Private Const N = 3\nPrivate Type Row3\n    vals(1 To N, 0 To 1) As Long\nEnd Type\n', ['Dim r As Row3', 'Main = UBound(r.vals, 2)']],
		] as const) {
			expect(errors(source(decl, ...lines)), lines.join(': ')).toEqual([]);
		}
	});

	it('reports a subscript or a dimension outside the bounds the Const gives', () => {
		for (const [decl, lines, span] of [
			[CONST_TYPE, ['Dim r As Row3', 'Main = r.vals(4)'], '4'],
			[CONST_TYPE, ['Dim r As Row3', 'Main = r.vals(0)'], '0'],
			['Private Const N = 3\nPrivate Type Row3\n    vals(N) As Long\nEnd Type\n', ['Dim r As Row3', 'Main = r.vals(4)'], '4'],
			['Private Enum Sz\n    szMax = 3\nEnd Enum\nPrivate Type Row3\n    vals(1 To szMax) As Long\nEnd Type\n', ['Dim r As Row3', 'Main = r.vals(4)'], '4'],
			['Private Const N = 2\nPrivate Type Row3\n    vals(N * 2) As Long\nEnd Type\n', ['Dim r As Row3', 'Main = r.vals(5)'], '5'],
			['Private Const N = 3\nPrivate Type Row3\n    vals(1 To N, 0 To 1) As Long\nEnd Type\n', ['Dim r As Row3', 'Main = r.vals(1, 2)'], '2'],
		] as const) {
			const src = source(decl, ...lines);
			expectDiagnostic(src, byCode(analyzeModule(src), 'array-subscript-out-of-bounds'), 'array-subscript-out-of-bounds', { span });
		}
		const dimension = source(CONST_TYPE, 'Dim r As Row3', 'Main = UBound(r.vals, 2)');
		expect(byCode(analyzeModule(dimension), 'array-subscript-out-of-bounds')).toHaveLength(1);
	});

	it('is fixed, so ReDim cannot resize it', () => {
		const src = source(CONST_TYPE, 'Dim r As Row3', 'ReDim r.vals(5)', 'Main = 1');
		expect(byCode(analyzeModule(src), 'fixed-array-redim')).toHaveLength(1);
	});

	it('stays fixed when another module declares the Const, though its bounds are not known', () => {
		const caller = source('Private Type Row3\n    vals(1 To M) As Long\nEnd Type\n', 'Dim r As Row3', 'Main = UBound(r.vals)');
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Module2', source: 'Option Explicit\nPublic Const M = 3\n' },
		], 'Module1');
		expect(diagnostics.filter((diag) => diag.severity === 'error').map((diag) => diag.code)).toEqual([]);
		const redim = source('Private Type Row3\n    vals(1 To M) As Long\nEnd Type\n', 'Dim r As Row3', 'ReDim r.vals(5)', 'Main = 1');
		const refused = analyzeProjectModule(redim, [{ moduleName: 'Module2', source: 'Option Explicit\nPublic Const M = 3\n' }], 'Module1');
		expect(byCode(refused, 'fixed-array-redim')).toHaveLength(1);
		// LSet between two Types takes one whose every member is of fixed size.
		const lset = source('Private Type A\n    vals(1 To M) As Long\nEnd Type\nPrivate Type B\n    x As Long\nEnd Type\n', 'Dim a1 As A, b1 As B', 'b1.x = 5', 'LSet a1 = b1', 'Main = a1.vals(1)');
		const copied = analyzeProjectModule(lset, [{ moduleName: 'Module2', source: 'Option Explicit\nPublic Const M = 3\n' }], 'Module1');
		expect(byCode(copied, 'lset-type-mismatch')).toEqual([]);
	});

	it('takes an implicit lower bound of 0', () => {
		const src = source('Private Const N = 3\nPrivate Type Row3\n    vals(N) As Long\nEnd Type\n', 'Dim r As Row3', 'r.vals(0) = 7', 'Main = r.vals(0)');
		expect(errors(src)).toEqual([]);
	});
});
