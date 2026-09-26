// Diagnostics tests: a Variant whose value is plain, used as another kind
// (issue #121). Each raising sample was measured in Excel 16.0 (build 20326,
// 2026-09-26); each quiet one runs there.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic, expectDiagnostics } from '../helpers/diagnostics';

const CODE = 'variant-value-misuse';

function wrap(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

describe('variant-value-misuse (issue #121)', () => {
	it('flags a member access on a Variant holding a number, a string or an array', () => {
		const cases: Array<[string, string]> = [
			['v = 5', 'the number 5'],
			['v = "abc"', 'the string "abc"'],
			['v = Array(1, 2)', 'an array from Array(...)'],
		];
		for (const [assignment, holds] of cases) {
			const src = wrap('Dim v As Variant', assignment, 'v.Foo', 'Main = 1');
			expectDiagnostic(src, analyzeModule(src), CODE, { span: 'v', message: [holds, "'424'"] });
		}
	});

	it('flags UBound and LBound of a Variant holding a scalar', () => {
		const src = wrap('Dim v As Variant', 'v = 5', 'Main = UBound(v)', 'Main = LBound(v, 1)');
		expectDiagnostics(src, analyzeModule(src), CODE, [
			{ span: 'v', message: ['not an array', "'13'"] },
			{ span: 'v', message: "'13'" },
		]);
	});

	it('flags an array Variant beside a scalar operator', () => {
		const cases: Array<[string, string]> = [
			['Main = v + 1', '+'],
			['Main = v - 1', '-'],
			['Main = v & "x"', '&'],
			['If v = 1 Then Main = 2', '='],
		];
		for (const [use, operator] of cases) {
			const src = wrap('Dim v As Variant', 'v = Array(1, 2)', use);
			expectDiagnostics(src, analyzeModule(src), CODE, [{ span: 'v', message: [`'${operator}'`, "'13'"] }]);
		}
	});

	it('stays quiet for indexing, UBound of an array, a typed variable and a reassigned one', () => {
		const src = wrap(
			'Dim v As Variant, w As Variant, l As Long',
			'v = Array(1, 2)',
			'Main = v(0) + 1',
			'Main = UBound(v)',
			'w = 5',
			'If Main Then w = Array(1)',
			'Main = UBound(w)',
			'l = 5',
			'Main = l + 1',
		);
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});
});
