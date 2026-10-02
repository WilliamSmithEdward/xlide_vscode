// Diagnostics tests: a label nothing jumps to below Exit Function does not
// hide what runs above it (issue #439, a regression from #421). Measured in
// Excel 16.0 (build 20326, 2026-10-01).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode } from '../helpers/diagnostics';

function source(use: string[], tail: string[]): string {
	return `Option Explicit\nPublic Function Main() As Variant\n    Dim o As Collection, a() As Long\n${[...use, 'Exit Function', ...tail].map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

const TAILS = [['L1:', 'Return'], ['L1:', 'Main = 2'], ['Bump:', 'Main = 3', 'Return'], ['EH:', 'Resume Next']];
const BLOCKS = (line: string): string[][] => [
	['If True Then', `    ${line}`, 'End If'],
	['Select Case 1', 'Case 1', `    ${line}`, 'End Select'],
	['With Application', `    ${line}`, 'End With'],
	[line],
];

describe('a label nothing jumps to below Exit Function', () => {
	it('leaves a Nothing object and an unallocated array above it reported', () => {
		for (const tail of TAILS) {
			for (const [line, code] of [['Main = o.Count', 'object-variable-not-set'], ['Main = a(0)', 'unallocated-dynamic-array-access']]) {
				for (const use of BLOCKS(line)) {
					expect(byCode(analyzeModule(source(use, tail)), code).length, `${use.join(' / ')} | ${tail.join(' / ')}`).toBe(1);
				}
			}
		}
	});
});
