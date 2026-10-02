// Diagnostics tests: financial and Format* built-ins with arguments the code
// proves wrong (issue #476). Each sample was measured through pyVBAharness on
// 2026-10-02 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const SETUP = [
	'Dim d As Double, l As Long, s As String, v As Variant, r As Double',
	'Dim da(2) As Double',
	'd = 1: l = 1: s = "abc": v = 5: r = -1',
	'da(0) = -1000: da(1) = 600: da(2) = 600',
];

function errors(expr: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n${[...SETUP, `Main = ${expr}`].map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
	return analyzeModule(src).filter((d) => d.severity === 'error').map((d) => /Run-time error '(\d+)'/.exec(d.message)?.[1] ?? d.code ?? '');
}

describe('financial and Format* arguments (issue #476)', () => {
	it('reports a rate of -1, text that is no number, a Tristate out of range, and Filter of a scalar', () => {
		const cases: Array<[string, string]> = [
			['NPV(-1, da)', '5'], ['NPV(r, da)', '5'], ['MIRR(da, -1, 0.1)', '5'], ['MIRR(da, 0.1, -1)', '5'], ['IRR(da, -1)', '5'],
			['FormatNumber("abc")', '13'], ['FormatNumber(s)', '13'], ['FormatCurrency("abc")', '13'], ['FormatPercent("abc")', '13'], ['FormatDateTime("abc")', '13'],
			['FormatNumber(1, 2, 5)', '5'], ['FormatNumber(1, 2, -3)', '5'], ['FormatNumber(1, 2, , , 5)', '5'],
			['Filter(v, "a")(0)', '13'],
		];
		for (const [expr, error] of cases) {
			expect(errors(expr), expr).toEqual([error]);
		}
	});

	it('stays quiet where the arguments are taken', () => {
		for (const expr of ['Pmt(-1, 12, 1000)', 'FV(-1, 12, -100)', 'FormatNumber("12.5")', 'FormatDateTime(1)', 'FormatNumber(Empty)', 'FormatNumber(Null)', 'FormatDateTime(Null)', 'FormatNumber(1, 2, -2)', 'NPV(0.1, da)']) {
			expect(errors(expr), expr).toEqual([]);
		}
	});
});
