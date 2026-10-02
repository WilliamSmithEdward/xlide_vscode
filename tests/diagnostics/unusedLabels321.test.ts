// Diagnostics tests: a label nothing jumps to, and line numbers, leave the
// state rules as they were (issue #321). Each sample was run through
// pyVBAharness on 2026-10-02 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(lines: string[]): string[] {
	const source = `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(source).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

const CASES: Array<[string, string[], string]> = [
	['an index past the end', ['Dim c As New Collection', 'c.Add 1', 'Main = c(5)'], 'collection-index-out-of-range'],
	['a key never added', ['Dim c As New Collection', 'c.Add 1, "a"', 'Main = c("nokey")'], 'collection-key-not-found'],
	['a key added twice', ['Dim c As New Collection', 'c.Add 1, "a"', 'c.Add 2, "a"'], 'collection-key-in-use'],
];

function numbered(lines: string[]): string[] {
	return lines.map((line, k) => (line.startsWith('Dim ') ? line : `${(k + 1) * 10} ${line}`));
}

describe('a label nothing jumps to (issue #321)', () => {
	it.each(CASES)('keeps %s reported', (_name, lines, code) => {
		expect(errors(lines)).toEqual([code]);
		expect(errors([...lines.slice(0, -1), 'L3:', lines[lines.length - 1]])).toEqual([code]);
		expect(errors(numbered(lines))).toEqual([code]);
	});

	it.each(CASES)('forgets %s at a label a GoTo names', (_name, lines) => {
		expect(errors([...lines.slice(0, -1), 'GoTo L2', 'L2:', lines[lines.length - 1]])).toEqual([]);
	});
});
