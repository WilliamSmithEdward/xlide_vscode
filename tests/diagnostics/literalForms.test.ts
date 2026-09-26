// Literal tokens the VBE refuses (issue #133). Each refused sample was
// measured in Excel 16.0 (build 20326, 2026-09-26) as a compile-time Syntax
// error; each accepted one compiles there.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

const SUFFIX = 'suffixed-literal-overflow';
const FLOAT = 'float-literal-overflow';
const DATE = 'date-literal-invalid';

function wrap(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

describe('type-suffixed and radix literals (issue #133)', () => {
	it('flags each refused suffixed literal', () => {
		const cases: Array<[string, string, string]> = [
			['2147483648&', SUFFIX, 'Long range'],
			['-2147483648&', SUFFIX, 'Long range'],
			['&H10000%', SUFFIX, 'four hex digits'],
			['&H100000000&', SUFFIX, 'eight hex digits'],
			['9223372036854775808^', SUFFIX, 'LongLong range'],
			['1E3%', SUFFIX, 'exponent'],
			['1.5%', SUFFIX, 'fractional'],
			['&H', SUFFIX, 'no digits'],
			['3.5E+38!', FLOAT, 'Single range'],
			['922337203685477.5808@', FLOAT, 'Currency range'],
			['922337203685478@', FLOAT, 'Currency range'],
			['1.8E+308#', FLOAT, 'Double range'],
			['1E+309', FLOAT, 'Double range'],
		];
		for (const [literal, code, message] of cases) {
			const src = wrap(`Main = ${literal}`);
			expectDiagnostic(src, analyzeModule(src), code, { message });
		}
	});

	it('accepts every literal at the edge of its range', () => {
		const src = wrap(
			'Dim v As Variant',
			'v = &H8000%', 'v = &HFFFF%', 'v = &O100000%', 'v = 2147483647&', 'v = &HFFFFFFFF&',
			'v = 3.402823E+38!', 'v = 922337203685477.5807@', 'v = 1.79769313486231E+308#',
			'v = 9223372036854775807^', 'v = 1E+308', 'v = 99999999999999999999', 'v = 3000000000&"x"',
			'Main = 1',
		);
		expect(byCode(analyzeModule(src), SUFFIX)).toHaveLength(0);
		expect(byCode(analyzeModule(src), FLOAT)).toHaveLength(0);
	});
});

describe('date literals (issue #133)', () => {
	it('flags a year past 9999, an impossible day, an impossible hour and an empty literal', () => {
		const cases: Array<[string, string]> = [
			['#1/1/10000#', 'year past 9999'],
			['#2/30/2000#', 'day 30'],
			['#1/0/2000#', 'day 0'],
			['#25:00#', 'hour 25'],
			['#1/1/2000 24:00:00#', 'hour 24'],
			['##', 'empty'],
		];
		for (const [literal, message] of cases) {
			const src = wrap(`Main = ${literal}`);
			expectDiagnostic(src, analyzeModule(src), DATE, { span: literal, message });
		}
	});

	it('accepts the edges and the day-first reading of a month past 12', () => {
		const src = wrap('Dim v As Variant', 'v = #12/31/9999#', 'v = #1/1/100#', 'v = #13/1/2000#', 'v = #13:00 PM#', 'v = #2/29/2000#', 'Main = 1');
		expect(byCode(analyzeModule(src), DATE)).toHaveLength(0);
	});
});
