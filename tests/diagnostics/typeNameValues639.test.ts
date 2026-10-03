// Diagnostics tests: a Type or Enum sharing its name with a value (issue
// #639). Measured on 2026-10-02 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeProjectModule } from './helpers';

const ENUM = 'Private Enum Zq\n    ZqA = 1\nEnd Enum\n';
const TYPE = 'Private Type Zq\n    v As Long\nEnd Type\n';
const FUNC = 'Private Function Zq() As Long\n    Zq = 2\nEnd Function\n';

function errors(decls: string, body: string, module2?: string): string[] {
	const src = `Option Explicit\n${decls}Function Main() As Variant\n    ${body}\nEnd Function\n`;
	const modules = [{ moduleName: 'Module1', source: src }, ...(module2 ? [{ moduleName: 'Module2', source: `Option Explicit\n${module2}` }] : [])];
	return analyzeProjectModule(src, modules, 'Module1').filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a Type or Enum sharing its name with a value (issue #639)', () => {
	it('reads the Function, variable or Const of the name, which runs', () => {
		expect(errors(ENUM + FUNC, 'Main = Zq')).toEqual([]);
		expect(errors(TYPE + FUNC, 'Main = Zq')).toEqual([]);
		expect(errors(TYPE, 'Main = Zq', 'Public Zq As Long\n')).toEqual([]);
		expect(errors(TYPE, 'Main = Zq', 'Public Const Zq As Long = 3\n')).toEqual([]);
		expect(errors(TYPE, 'Zq = 5\n    Main = Zq', 'Public Zq As Long\n')).toEqual([]);
		expect(errors(TYPE + 'Private Zq As Long\n', 'Main = Zq')).toEqual([]);
	});

	it('still reports the Enum or Type read as a value', () => {
		expect(errors(ENUM, 'Main = Zq', 'Public Function Zq() As Long\n    Zq = 2\nEnd Function\n')).toEqual(['malformed-statement']);
		expect(errors(ENUM + 'Private Zq As Long\n', 'Main = Zq')).toEqual(['malformed-statement']);
		expect(errors(TYPE, 'Main = Zq')).toEqual(['undeclared-variable']);
	});

	it('reports TypeName of a Type or Enum, and a Sub beside a Type', () => {
		expect(errors(TYPE, 'Main = TypeName(Zq)')).toEqual(['undeclared-variable']);
		expect(errors(ENUM, 'Main = TypeName(Zq)')).toEqual(['malformed-statement']);
		expect(errors(ENUM + 'Private Zq As Long\n', 'Main = TypeName(Zq)')).toEqual(['malformed-statement']);
		expect(errors(TYPE + 'Private Sub Zq()\nEnd Sub\n', 'Main = Zq')).toEqual(['sub-used-as-value']);
	});

	it('reports two Enums or two Types of one name', () => {
		expect(errors(ENUM + 'Private Enum Zq\n    ZqB = 2\nEnd Enum\n', 'Main = 1')).toEqual(['type-enum-name-conflict']);
		expect(errors(TYPE + 'Private Type Zq\n    w As Long\nEnd Type\n', 'Main = 1')).toEqual(['type-enum-name-conflict']);
	});
});
