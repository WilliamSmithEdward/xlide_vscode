// Diagnostics tests: every statement's file modes, and an open file removed,
// copied or renamed (issue #419). Measured in Excel 16.0 64-bit (build
// 20430, 2026-10-02) through pyVBAharness.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(...lines: string[]): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    Dim p As String, s As String, n As Long\n    p = "C:\\t.txt"\n${lines.map((line) => `    ${line}`).join('\n')}\n    Close #1\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => `${diag.code}: ${diag.message}`);
}

// Each statement and the modes it runs in; every other mode raises 54.
const RUNS: ReadonlyArray<[string, readonly string[]]> = [
	['Print #1, "x"', ['Output', 'Append']],
	['Write #1, "x"', ['Output', 'Append']],
	['Input #1, s', ['Input', 'Binary']],
	['Line Input #1, s', ['Input', 'Binary']],
	['s = Input(1, #1)', ['Input', 'Binary']],
	['Get #1, , n', ['Binary', 'Random']],
	['Put #1, , n', ['Binary', 'Random']],
];

describe('a file statement in each mode', () => {
	it('raises 54 outside the modes it works in', () => {
		for (const [statement, modes] of RUNS) {
			for (const mode of ['Input', 'Output', 'Append', 'Binary', 'Random']) {
				const found = errors(`Open p For ${mode} As #1`, statement);
				expect(found, `${statement} For ${mode}`).toEqual(modes.includes(mode) ? [] : [expect.stringMatching(/^file-mode-mismatch: .*'54'/)]);
			}
		}
	});
});

describe('an open file removed, copied or renamed', () => {
	it('raises 55', () => {
		for (const lines of [
			['Open p For Input As #1', 'Kill p'], ['Open p For Binary As #1', 'Kill p'], ['Open p For Append As #1', 'FileCopy p, p & ".bak"'],
			['Open p For Output As #1', 'FileCopy p, p & ".bak"'], ['Open p For Random As #1', 'FileCopy p, p & ".bak"'],
			['Open p For Input As #1', 'Name p As p & ".x"'],
		]) {
			expect(errors(...lines), lines.join(' / ')).toEqual([expect.stringMatching(/^file-already-open: .*'55'/)]);
		}
	});

	it('runs once closed, or copied while open For Input', () => {
		expect(errors('Open p For Input As #1', 'Close #1', 'Kill p')).toEqual([]);
		expect(errors('Open p For Input As #1', 'FileCopy p, p & ".bak"')).toEqual([]);
		expect(errors('Open p For Input As #1', 'p = "C:\\u.txt"', 'Kill p')).toEqual([]);
	});
});
