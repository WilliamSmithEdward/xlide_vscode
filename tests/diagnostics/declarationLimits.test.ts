// Declaration limits the VBE enforces (issue #210). Measured in 64-bit Excel
// 16.0: the parameter limit in a class module, Event and Declare parameter
// counts, line numbers outside the Long range, and an Enum member that is a
// string.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode } from '../helpers/diagnostics';

const params = (count: number): string => Array.from({ length: count }, (_, i) => `p${i}&`).join(', ');

describe('too-many-parameters (issue #210)', () => {
	const inClass = (src: string) => byCode(analyzeModule(src, { moduleKind: 'class' }), 'too-many-parameters');
	const inStandard = (src: string) => byCode(analyzeModule(src, { moduleKind: 'standard' }), 'too-many-parameters');

	it.each([
		['Public Sub', (n: number) => `Public Sub M(${params(n)})\nEnd Sub\n`],
		['Friend Sub', (n: number) => `Friend Sub M(${params(n)})\nEnd Sub\n`],
		['Function', (n: number) => `Public Function M(${params(n)}) As Long\nEnd Function\n`],
		['Property Get', (n: number) => `Public Property Get M(${params(n)}) As Long\nEnd Property\n`],
		['Property Let, its value counted', (n: number) => `Public Property Let M(${params(n - 1)}, ByVal v As Long)\nEnd Property\n`],
		['Event', (n: number) => `Public Event E(${params(n)})\n`],
		['Private Declare', (n: number) => `Private Declare PtrSafe Sub Sleep2 Lib "kernel32" Alias "Sleep" (${params(n)})\n`],
		['Sub with a ParamArray', (n: number) => `Public Sub M(${params(n - 1)}, ParamArray rest())\nEnd Sub\n`],
	])('allows 59 in a class module and refuses 60: %s', (_name, source) => {
		expect(inClass(source(59))).toHaveLength(0);
		const hits = inClass(source(60));
		expect(hits).toHaveLength(1);
		expect(hits[0].message).toContain('In a class module');
		expect(hits[0].message).toContain('at most 59 parameters');
	});

	it('allows 60 in a standard module, a Declare included, and refuses 61', () => {
		const declare = (n: number) => `Private Declare PtrSafe Sub Sleep2 Lib "kernel32" Alias "Sleep" (${params(n)})\n`;
		expect(inStandard(`Public Sub M(${params(60)})\nEnd Sub\n`)).toHaveLength(0);
		expect(inStandard(declare(60))).toHaveLength(0);
		expect(inStandard(declare(61))[0].message).toBe("A Declare may have at most 60 parameters; 'Sleep2' declares 61.");
	});
});

describe('invalid-line-number (issue #210)', () => {
	const lines = (line: string) => byCode(analyzeModule(`Function Main() As Variant\n${line}\nEnd Function\n`), 'invalid-line-number');

	it('refuses a line number past the Long range, or a negative one', () => {
		expect(lines('2147483648 Main = 2')).toHaveLength(1);
		expect(lines('99999999999 Main = 2')).toHaveLength(1);
		expect(lines('-1 Main = 2')).toHaveLength(1);
	});

	it('accepts 0 and 2147483647', () => {
		expect(lines('2147483647 Main = 2')).toHaveLength(0);
		expect(lines('0 Main = 2')).toHaveLength(0);
	});
});

describe('enum-member-type-mismatch (issue #210)', () => {
	const members = (setup: string, value: string) =>
		byCode(analyzeModule(`${setup}Private Enum E\n    a = ${value}\nEnd Enum\n`), 'enum-member-type-mismatch');

	it('refuses a string no locale reads as a number', () => {
		expect(members('', '"x"')).toHaveLength(1);
		expect(members('', '"a" & "b"')).toHaveLength(1);
		expect(members('Private Const S As String = "x"\n', 'S')).toHaveLength(1);
		// Refused too, though CLng("True") runs.
		expect(members('', '"True"')).toHaveLength(1);
		expect(members('', '""')).toHaveLength(1);
	});

	it('accepts a numeric string and the values VBA converts', () => {
		for (const value of ['"1"', '1.5', 'True', '#1/1/2000#', '"&H10"']) {
			expect(members('', value), value).toHaveLength(0);
		}
	});
});
