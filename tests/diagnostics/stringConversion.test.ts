// What VBA makes of a string it converts (issue #188). Every verdict below
// was measured in Excel 16.0 (build 20326, en-US, 2026-09-29) with
// `x = "..."` into a typed variable. A string another locale reads as a
// number ("1 000", "5x") is left alone even where en-US refuses it.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import {
	isInvalidBooleanString,
	isInvalidDateString,
	numericStringVerdict,
} from '../../src/analyzer/diagnostics/stringConversion';
import { byCode } from '../helpers/diagnostics';

describe('numericStringVerdict (issue #188)', () => {
	it('reads hex, octal and plain numbers, and every form en-US converts', () => {
		expect(numericStringVerdict('&HFF')).toEqual({ kind: 'number', value: 255 });
		expect(numericStringVerdict('&hff')).toEqual({ kind: 'number', value: 255 });
		expect(numericStringVerdict('&HFFFF')).toEqual({ kind: 'number', value: -1 });
		expect(numericStringVerdict('&H80000000')).toEqual({ kind: 'number', value: -2147483648 });
		expect(numericStringVerdict('&17')).toEqual({ kind: 'number', value: 15 });
		expect(numericStringVerdict(' 5 ')).toEqual({ kind: 'number', value: 5 });
		expect(numericStringVerdict('(5)')).toEqual({ kind: 'number', value: -5 });
		expect(numericStringVerdict('5-')).toEqual({ kind: 'number', value: -5 });
		expect(numericStringVerdict('1e3')).toEqual({ kind: 'number', value: 1000 });
		for (const text of ['$5', '5 $', '1,000', '2.5', '0.0', '+5']) {
			expect(numericStringVerdict(text).kind, text).toBe('number');
		}
		// Separators and currency are the locale's: no value is claimed.
		expect(numericStringVerdict('1,000')).toEqual({ kind: 'number' });
		expect(numericStringVerdict('$5')).toEqual({ kind: 'number' });
	});

	it('refuses what no locale converts', () => {
		for (const text of ['', '   ', 'abc', '.', '$', '5%', '-&H10', '4x2', '1/2/2020', '12:30']) {
			expect(numericStringVerdict(text).kind, JSON.stringify(text)).toBe('invalid');
		}
		// Refused in en-US, read by another locale.
		for (const text of ['1 000', '5x']) {
			expect(numericStringVerdict(text).kind, text).toBe('number');
		}
	});

	it('reads Boolean and Date strings the way the VBE does', () => {
		for (const text of ['True', 'true', 'TRUE', 'False', '5', '0.0', '$5', '(5)', '&HFF', ' 1 ']) {
			expect(isInvalidBooleanString(text), text).toBe(false);
		}
		for (const text of [' True ', 'yes', '', '4x2', '5%']) {
			expect(isInvalidBooleanString(text), JSON.stringify(text)).toBe(true);
		}
		for (const text of ['5', '1/2/2020', 'Jan 1', '12:30', '$5']) {
			expect(isInvalidDateString(text), text).toBe(false);
		}
		for (const text of ['5%', '.', '', 'True', '-&H10', 'May', 'Monday', 'abc']) {
			expect(isInvalidDateString(text), JSON.stringify(text)).toBe(true);
		}
	});
});

describe('string literals into typed variables (issue #188)', () => {
	const into = (type: string, literal: string): string[] => {
		const src = `Option Explicit\nSub Main()\n    Dim x As ${type}\n    x = ${literal}\nEnd Sub\n`;
		return byCode(analyzeModule(src), 'assignment-type-mismatch').map((hit) => hit.message);
	};

	it('stays quiet on the strings VBA converts', () => {
		for (const [type, literal] of [['Long', '"&HFF"'], ['Integer', '"&HFFFF"'], ['Byte', '"&HFF"'], ['Boolean', '"5"'], ['Boolean', '"2.5"'], ['Date', '"$5"'], ['Long', '"1,000"']]) {
			expect(into(type, literal), `${type} <- ${literal}`).toEqual([]);
		}
	});

	it('reports the strings that raise 13, and the ones that overflow with 6', () => {
		for (const [type, literal] of [['Long', '"4x2"'], ['Long', '"5%"'], ['Long', '""'], ['Date', '"True"'], ['Date', '"May"'], ['Boolean', '" True "']]) {
			const messages = into(type, literal);
			expect(messages, `${type} <- ${literal}`).toHaveLength(1);
			expect(messages[0]).toContain("'13'");
		}
		for (const [type, literal] of [['Integer', '"&H10000"'], ['Byte', '"&H100"'], ['Integer', '"40000"']]) {
			const messages = into(type, literal);
			expect(messages, `${type} <- ${literal}`).toHaveLength(1);
			expect(messages[0]).toContain("'6': Overflow");
		}
	});

	it('reports an empty or percent string in arithmetic and a comparison with a number', () => {
		const src = 'Option Explicit\nFunction Main() As Variant\n    Dim n As Long\n    Main = "" + 1\n    Main = "5%" * 2\n    If n = "" Then Main = 0\nEnd Function\n';
		expect(byCode(analyzeModule(src), 'string-arithmetic-coercion')).toHaveLength(3);
	});
});
