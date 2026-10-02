// Diagnostics tests: under On Error Resume Next, a Resume with no error
// pending raises 20, which is skipped, so the next line runs (issue #446).
// Measured in Excel 16.0 (build 20326, 2026-10-02).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode } from '../helpers/diagnostics';

const CODE = 'unreachable-code';

function source(...lines: string[]): string {
	return `Option Explicit\nPublic Function Main() As Variant\n    Dim n As Long\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

describe('a Resume under On Error Resume Next', () => {
	it('does not end the line', () => {
		for (const lines of [
			['On Error GoTo H', 'On Error Resume Next', 'Resume Next', 'n = n + 1', 'Exit Function', 'H:', 'Main = "h"'],
			['On Error Resume Next', 'Resume', 'Main = 1'],
			['On Error Resume Next', 'Resume L1', 'Main = 1', 'L1:', 'Main = Main + 1'],
		]) {
			expect(byCode(analyzeModule(source(...lines)), CODE), lines.join(' / ')).toEqual([]);
		}
	});

	it('still ends the line without it', () => {
		const src = source('On Error GoTo H', 'Exit Function', 'H:', 'Resume Next', 'Main = 2');
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(1);
	});
});
