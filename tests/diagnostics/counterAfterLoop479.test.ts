// Diagnostics tests: a loop counter used after its loop ends (issue #479).
// Each raising sample was measured through pyVBAharness on 2026-10-02 in
// Excel 16.0 (build 20430); each quiet one runs there.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode } from '../helpers/diagnostics';

const BOUNDS = 'array-subscript-out-of-bounds';

function wrap(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n    Dim a(1 To 5) As Long, i As Long\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

describe('a loop counter after its loop (issue #479)', () => {
	it('knows the counter past the end of a For or a Do loop', () => {
		const loops = [
			['For i = 1 To 5', '    a(i) = i', 'Next'],
			['For i = 1 To 5 Step 2', '    a(i) = i', 'Next'],
			['For i = 5 To 1 Step -1', '    a(i) = i', 'Next'],
			['i = 1', 'Do While i <= 5', '    a(i) = i', '    i = i + 1', 'Loop'],
			['i = 5', 'Do Until i < 1', '    a(i) = i', '    i = i - 1', 'Loop'],
			['i = 1', 'Do', '    a(i) = i', '    i = i + 1', 'Loop While i <= 5'],
			['i = 1', 'While i <= 5', '    i = i + 1', 'Wend'],
		];
		for (const lines of loops) {
			expect(byCode(analyzeModule(wrap(...lines, 'Main = a(i)')), BOUNDS), lines.join(' / ')).toHaveLength(1);
		}
	});

	it('stays quiet where the loop may leave early or ends inside the bounds', () => {
		const loops = [
			['For i = 1 To 5', '    If a(i) = 0 Then Exit For', 'Next'],
			['For i = 1 To 4', '    a(i) = i', 'Next'],
			['i = 1', 'Do While i <= 4', '    i = i + 1', 'Loop'],
			['i = 1', 'Do While i <= 5', '    If a(i) = 0 Then Exit Do', '    i = i + 1', 'Loop'],
		];
		for (const lines of loops) {
			expect(byCode(analyzeModule(wrap(...lines, 'Main = a(i)')), BOUNDS), lines.join(' / ')).toHaveLength(0);
		}
	});

	it('reports a Do loop whose counter overflows before the test can end it', () => {
		const src = 'Option Explicit\nFunction Main() As Variant\n    Dim i As Integer\n    Do Until i < 0\n        i = i + 1000\n    Loop\nEnd Function\n';
		expect(byCode(analyzeModule(src), 'arithmetic-overflow')).toHaveLength(1);
		const exits = 'Option Explicit\nFunction Main() As Variant\n    Dim i As Long\n    Do Until i < 0\n        i = i + 1000\n        If i > 40000 Then Exit Do\n    Loop\n    Main = i\nEnd Function\n';
		expect(byCode(analyzeModule(exits), 'arithmetic-overflow')).toHaveLength(0);
	});
});
