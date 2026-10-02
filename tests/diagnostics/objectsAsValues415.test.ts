// Diagnostics tests: object variables read as values (issue #415).
// Measured in Excel 16.0 64-bit (build 20430, 2026-10-02) through
// pyVBAharness.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(...lines: string[]): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
	return analyzeModule(src, { host: 'excel' }).filter((diag) => diag.severity === 'error').map((diag) => `${diag.code}: ${diag.message}`);
}

describe('an object never set, read as a value', () => {
	it('raises 91 through CStr, Len or an index', () => {
		for (const type of ['Range', 'Object']) {
			for (const use of ['Main = CStr(x)', 'Main = Len(x)', 'Main = x(1)']) {
				expect(errors(`Dim x As ${type}`, use), `${type}: ${use}`).toEqual([expect.stringMatching(/^object-variable-not-set: .*'91'/)]);
			}
		}
		expect(errors('Dim x As Application', 'Main = CStr(x)')).toEqual([expect.stringMatching(/^object-variable-not-set: /)]);
		expect(errors('Dim x As Collection', 'Main = x(1)')).toEqual([expect.stringMatching(/^object-variable-not-set: /)]);
	});

	it('is 438 or 91 for a type with no default member', () => {
		for (const type of ['Worksheet', 'Workbook', 'Font']) {
			for (const use of ['Main = CStr(x)', 'Main = Len(x)', 'Main = x(1)', 'If x Then Main = 1']) {
				expect(errors(`Dim x As ${type}`, use), `${type}: ${use}`).toEqual([expect.stringMatching(/^object-default-value: .*'438'.*'91' while it is Nothing/)]);
			}
		}
	});
});

describe('a set object with no default member', () => {
	it('raises 438 in a condition, an index, CStr and Len', () => {
		for (const [type, value] of [['Worksheet', 'ActiveSheet'], ['Workbook', 'ThisWorkbook'], ['Font', 'Range("A1").Font']]) {
			for (const use of ['If x Then Main = 1', 'Main = x(1)', 'Main = CStr(x)', 'Main = Len(x)']) {
				expect(errors(`Dim x As ${type}`, `Set x = ${value}`, use), `${type}: ${use}`).toEqual([expect.stringMatching(/^object-default-value: .*'438'/)]);
			}
		}
	});
});

describe('an Object holding a Collection', () => {
	it('raises 450 read as a value and 438 for a Let', () => {
		for (const use of ['Main = x', 'Main = x + 1', 'Main = x & "a"', 'If x Then Main = 1', 'Main = CStr(x)', 'Main = Len(x)', 'Main = (x = 0)']) {
			expect(errors('Dim x As Object', 'Set x = New Collection', use), use).toEqual([expect.stringMatching(/^object-default-value: .*'450'/)]);
		}
		expect(errors('Dim x As Object', 'Set x = New Collection', 'x = 5')).toEqual([expect.stringMatching(/^object-default-value: .*'438'/)]);
	});

	it('is quiet once it holds something else, and for IsNumeric', () => {
		expect(errors('Dim x As Object', 'Set x = New Collection', 'Set x = Range("A1")', 'Main = x + 1')).toEqual([]);
		expect(errors('Dim x As Object', 'Set x = New Collection', 'Main = IsNumeric(x)')).toEqual([]);
		expect(errors('Dim x As Range', 'Set x = Range("A1")', 'Main = CStr(x)')).toEqual([]);
	});
});
