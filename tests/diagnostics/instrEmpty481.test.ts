// Diagnostics tests: InStr with Start 0 and an empty string (issue #481).
// Each sample was measured through pyVBAharness on 2026-10-02 in Excel
// 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode } from '../helpers/diagnostics';

const VALUE = 'runtime-argument-value';

function wrap(expr: string): string {
	return `Option Explicit\nFunction Main() As Variant\n    Dim s As String, t As String\n    t = "abc"\n    Main = ${expr} & s & t\nEnd Function\n`;
}

describe('InStr with an empty string (issue #481)', () => {
	it('stays quiet when either string is empty, whatever Start and Compare are', () => {
		for (const expr of ['InStr(0, "abc", "")', 'InStr(0, "", "b")', 'InStr(0, "", "")', 'InStr(-1, "abc", "")', 'InStr(0, s, "b")', 'InStr(0, "abc", vbNullString)', 'InStr(0, "abc", "", 5)']) {
			expect(byCode(analyzeModule(wrap(expr)), VALUE), expr).toHaveLength(0);
		}
	});

	it('still reports InStrB, InStrRev, Mid, Left and InStr with both strings there', () => {
		for (const expr of ['InStr(0, "abc", "b")', 'InStr(0, t, "b")', 'InStrB(0, "abc", "")', 'InStrRev("abc", "", 0)', 'Mid("", 0)', 'Left("", -1)']) {
			expect(byCode(analyzeModule(wrap(expr)), VALUE), expr).toHaveLength(1);
		}
	});
});
