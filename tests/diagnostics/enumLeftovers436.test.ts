// Diagnostics tests: Enum leftovers (issue #436). Measured in Excel 16.0
// 64-bit (build 20430, 2026-10-02) through pyVBAharness.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(head: string, ...lines: string[]): string[] {
	const src = `Option Explicit\n${head}Function Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => `${diag.code}: ${diag.message}`);
}

const ENUM = (member: string): string => `Private Enum E\n    ${member}\nEnd Enum\n`;

describe('an Enum member named like another module declaration', () => {
	it('is Ambiguous name detected', () => {
		for (const head of [
			ENUM('Helper') + 'Private Sub Helper()\nEnd Sub\n',
			ENUM('Helper') + 'Public Sub Helper()\nEnd Sub\n',
			ENUM('Helper') + 'Private Function Helper() As Long\nEnd Function\n',
			ENUM('zz') + 'Private zz As Long\n',
			ENUM('kk') + 'Private Const kk = 1\n',
		]) {
			expect(errors(head, 'Main = 1'), head).toEqual([expect.stringMatching(/^duplicate-procedure: Ambiguous name detected/)]);
		}
	});

	it('compiles beside a local of the same name', () => {
		expect(errors(ENUM('qq'), 'Dim qq As Long', 'Main = qq')).toEqual([]);
	});
});

describe('an Enum type', () => {
	it('has no value of its own', () => {
		expect(errors(ENUM('a = 3'), 'Main = E')).toEqual([expect.stringMatching(/^malformed-statement: .*not enum type/)]);
		expect(errors(ENUM('a = 3'), 'Main = E.a')).toEqual([]);
	});

	it('types a variable as a Long', () => {
		expect(errors(ENUM('a = 3'), 'Dim x As E', 'x = "abc"')).toEqual([expect.stringMatching(/^assignment-type-mismatch: .*'13'/)]);
		for (const value of ['3000000000#', '2147483648#']) {
			expect(errors(ENUM('a = 3'), 'Dim x As E', `x = ${value}`), value).toEqual([expect.stringMatching(/^assignment-type-mismatch: .*'6'/)]);
		}
		for (const value of ['99', '"5"', '-2147483648#']) {
			expect(errors(ENUM('a = 3'), 'Dim x As E', `x = ${value}`), value).toEqual([]);
		}
	});
});
