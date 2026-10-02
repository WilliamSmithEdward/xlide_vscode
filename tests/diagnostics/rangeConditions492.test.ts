// Diagnostics tests: multi-cell ranges in conditions, and the block If line
// read as a statement (issue #492). Each sample was measured through
// pyVBAharness on 2026-10-02 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(...lines: string[]): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
	return analyzeModule(src).filter((d) => d.severity === 'error').map((d) => d.code ?? '');
}

const WS = ['Dim ws As Worksheet', 'Set ws = ActiveSheet'];

describe('multi-cell ranges in conditions (issue #492)', () => {
	it('reports a block of cells read as a condition or a value', () => {
		const bodies = [
			[...WS, 'If ws.Range("A1:A2") Then', '    Main = 1', 'End If'],
			[...WS, 'If ws.Range("A1:A2").Value = 1 Then', '    Main = 1', 'End If'],
			[...WS, 'If False Then', '    Main = 3', 'ElseIf ws.Range("A1:A2") Then', '    Main = 1', 'End If'],
			[...WS, 'If ws.Range("A1:A2") Then Main = 1 Else Main = 2'],
			[...WS, 'Main = ws.Range(ws.Cells(1, 1), ws.Cells(2, 1)).Value + 1'],
			[...WS, 'Do While ws.Range("A1:A2")', '    Exit Do', 'Loop'],
		];
		for (const body of bodies) {
			expect(errors(...body), body.join(' / ')).toEqual(['multi-cell-range-as-scalar']);
		}
		expect(errors(...WS, 'Dim x As Variant', 'x = ws.Range("A1:A3").Value', 'Main = Join(x, ",")')).toEqual(['runtime-argument-value']);
	});

	it('reads a block If line as the value rules read a statement', () => {
		const cases: Array<[string[], string]> = [
			[['Dim d As Long', 'If 10 / d > 1 Then', '    Main = 1', 'End If'], 'division-by-zero'],
			[['If Mid("abc", 0) = "" Then', '    Main = 1', 'End If'], 'runtime-argument-value'],
			[['If CLng("abc") = 1 Then', '    Main = 1', 'End If'], 'runtime-conversion-value'],
			[['If Cells(0, 1).Value = 1 Then', '    Main = 1', 'End If'], 'host-argument-out-of-range'],
		];
		for (const [body, code] of cases) {
			expect(errors(...body), body.join(' / ')).toEqual([code]);
		}
	});

	it('stays quiet on one cell, and on a condition that holds', () => {
		for (const body of [
			[...WS, 'If ws.Range("A1").Value = 0 Then', '    Main = 1', 'End If'],
			[...WS, 'Main = ws.Range(ws.Cells(1, 1), ws.Cells(1, 1)).Value + 1'],
			[...WS, 'Dim x As Variant', 'x = ws.Range("A1").Value', 'Main = x'],
			['Dim d As Long', 'd = 2', 'If 10 / d > 1 Then', '    Main = 1', 'End If'],
			[...WS, 'If ws.Range("A1:A2").Cells(1, 1).Value = 0 Then', '    Main = 1', 'End If'],
		]) {
			expect(errors(...body), body.join(' / ')).toEqual([]);
		}
	});
});
