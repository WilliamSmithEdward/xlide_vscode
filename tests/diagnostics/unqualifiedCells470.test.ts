// Diagnostics tests: an unqualified Cells, Range or Rows inside a Range on a
// sheet the code knows is not active (issue #470). Each sample was measured
// through pyVBAharness on 2026-10-02 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode } from '../helpers/diagnostics';

const RANGE = 'host-argument-out-of-range';
const SETUP = ['Dim w1 As Worksheet, w2 As Worksheet, r As Range', 'Set w1 = ActiveSheet', 'Set w2 = Worksheets.Add'];

function hits(...lines: string[]) {
	const src = `Option Explicit\nFunction Main() As Variant\n${[...SETUP, ...lines].map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
	return byCode(analyzeModule(src), RANGE);
}

describe('unqualified corners of another sheet\'s Range (issue #470)', () => {
	it('does not infer active-sheet identity from an earlier Add', () => {
		for (const lines of [
			['Set r = w1.Range(Cells(1, 1), Cells(2, 2))'],
			['Set r = w1.Range(Range("A1"), Range("B2"))'],
			['Set r = w1.Range(Rows(1), Rows(2))'],
			['With w1', '    Set r = .Range(Cells(1, 1), .Cells(2, 2))', 'End With'],
		]) {
			const found = hits(...lines);
			expect(found, lines.join(' / ')).toHaveLength(0);
		}
	});

	it('stays quiet on the active sheet, every part qualified, after Activate, and on an address', () => {
		for (const lines of [
			['Set r = w2.Range(Cells(1, 1), Cells(2, 2))'],
			['Set r = w1.Range(w1.Cells(1, 1), w1.Cells(2, 2))'],
			['With w1', '    Set r = .Range(.Cells(1, 1), .Cells(2, 2))', 'End With'],
			['w1.Activate', 'Set r = w1.Range(Cells(1, 1), Cells(2, 2))'],
			['Set r = w1.Range(Cells(1, 1).Address)'],
		]) {
			expect(hits(...lines), lines.join(' / ')).toHaveLength(0);
		}
	});
});
