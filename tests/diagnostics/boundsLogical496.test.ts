// Diagnostics tests: And, Or, Xor, Not, Mod and \ in an array bound (issue
// #496). Each sample was compiled through pyVBAharness on 2026-10-02 in
// Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { evaluateIntegerConstantExpression } from '../../src/analyzer/constants/integerConstantExpression';
import { byCode } from '../helpers/diagnostics';

const BOUNDS = 'array-declaration-impossible-bounds';

function module(constant: string, dim: string): string {
	return `Option Explicit\n${constant ? `${constant}\n` : ''}Function Main() As Variant\n    ${dim}\n    Main = UBound(m)\nEnd Function\n`;
}

describe('logical operators and Mod in a bound (issue #496)', () => {
	it('folds them at VBA precedence', () => {
		const none = new Map<string, number>();
		const lookup = { get: (name: string) => none.get(name) };
		expect(evaluateIntegerConstantExpression('15 And 255', lookup)).toBe(15);
		expect(evaluateIntegerConstantExpression('50 Mod 7 + 10', lookup)).toBe(11);
		expect(evaluateIntegerConstantExpression('1 Or 8', lookup)).toBe(9);
		expect(evaluateIntegerConstantExpression('6 Xor 3', lookup)).toBe(5);
		expect(evaluateIntegerConstantExpression('Not 0', lookup)).toBe(-1);
		expect(evaluateIntegerConstantExpression('10 \\ 20', lookup)).toBe(0);
		expect(evaluateIntegerConstantExpression('7 \\ 2 * 3', lookup)).toBe(1);
		expect(evaluateIntegerConstantExpression('5 Mod 0', lookup)).toBeUndefined();
	});

	it('reports a lower bound past the upper', () => {
		const cases: Array<[string, string]> = [
			['Const K0 As Long = 15 And 255', 'Dim m(K0 To 3) As Long'],
			['Const K0 = 50 Mod 7 + 10', 'Dim m(K0 To 3) As Long'],
			['', 'Dim m(15 And 255 To 3) As Long'],
			['Const K0 = 1 Or 8', 'Dim m(K0 To 3) As Long'],
			['', 'Dim m(1 To 10 \\ 20) As Long'],
			['Const K0 = 6 Xor 3', 'Dim m(K0 To 3) As Long'],
			['', 'Dim m(1 To Not 0) As Long'],
		];
		for (const [constant, dim] of cases) {
			expect(byCode(analyzeModule(module(constant, dim)), BOUNDS), `${constant} / ${dim}`).toHaveLength(1);
		}
	});

	it('stays quiet on bounds that hold', () => {
		const cases: Array<[string, string]> = [
			['Const K0 = 15 And 255', 'Dim m(3 To K0) As Long'],
			['Const K0 = 50 Mod 7', 'Dim m(K0 To 3) As Long'],
			['', 'Dim m(1 To 20 \\ 10) As Long'],
			['', 'Dim m(Not -2 To 3) As Long'],
		];
		for (const [constant, dim] of cases) {
			expect(byCode(analyzeModule(module(constant, dim)), BOUNDS), `${constant} / ${dim}`).toHaveLength(0);
		}
	});
});
