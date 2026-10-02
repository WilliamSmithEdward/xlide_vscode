// Diagnostics tests: On n GoTo and On n GoSub with an index outside 0 to 255,
// or text (issue #499). Each sample was measured through pyVBAharness on
// 2026-10-02 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode } from '../helpers/diagnostics';

const VALUE = 'runtime-argument-value';
const TAIL = ['Main = 9', 'Exit Function', 'L1:', 'Main = 1', 'Exit Function', 'L2:', 'Main = 2', 'Exit Function', 'S1:', 'Main = 7', 'Return'];

function wrap(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n    Dim n As Long, d As Double, s As String\n${[...lines, ...TAIL].map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

describe('On n GoTo and On n GoSub index (issue #499)', () => {
	it('reports an index outside 0 to 255, and text', () => {
		const cases: Array<[string[], string]> = [
			[['On -1 GoTo L1'], "'5'"],
			[['On 300 GoTo L1'], "'5'"],
			[['n = 256', 'On n GoTo L1, L2'], "'5'"],
			[['n = -1', 'On n GoTo L1, L2'], "'5'"],
			[['n = 256', 'On n GoSub S1'], "'5'"],
			[['s = "x"', 'On s GoTo L1'], "'13'"],
		];
		for (const [lines, error] of cases) {
			const hits = byCode(analyzeModule(wrap(...lines)), VALUE);
			expect(hits, lines.join(' / ')).toHaveLength(1);
			expect(hits[0].message).toContain(error);
		}
	});

	it('stays quiet on an index that falls through or jumps', () => {
		for (const lines of [['n = 0', 'On n GoTo L1, L2'], ['n = 3', 'On n GoTo L1, L2'], ['n = 255', 'On n GoTo L1, L2'], ['d = 1.5', 'On d GoTo L1, L2'], ['n = 1', 'On n GoSub S1'], ['s = "2"', 'On s GoTo L1, L2'], ['On Error GoTo L1']]) {
			expect(byCode(analyzeModule(wrap(...lines)), VALUE), lines.join(' / ')).toHaveLength(0);
		}
	});
});
