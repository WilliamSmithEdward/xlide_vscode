// Diagnostics tests: Const leftovers after #458 (issue #494). Each sample
// was compiled through pyVBAharness on 2026-10-02 in Excel 16.0 (build
// 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(constant: string): string[] {
	const src = `Option Explicit\n${constant}\nFunction Main() As Variant\n    Main = C\nEnd Function\n`;
	return analyzeModule(src).filter((d) => d.severity === 'error').map((d) => d.code ?? '');
}

describe('Const leftovers after #458 (issue #494)', () => {
	it('reports a String that is no number under a logical operator', () => {
		for (const constant of ['Private Const C = "" And 255', 'Private Const C = 1 Or ""', 'Private Const C = "" Imp 7', 'Private Const C = "abc" Eqv 3', 'Private Const C = "abc" Xor 1']) {
			expect(errors(constant), constant).toEqual(['const-evaluation-error']);
		}
	});

	it('reads a numeric String beside a number, and a Currency exactly', () => {
		for (const constant of [
			'Private Const C As Byte = 1 - "1E3"',
			'Private Const C = 922337203685477.5807@ + 0.0001@',
			'Private Const C = 922337203685477.5807@ + "2"',
			'Private Const C As Integer = "2" * 20000',
		]) {
			expect(errors(constant), constant).toEqual(['const-overflow']);
		}
	});

	it('stays quiet on what compiles', () => {
		for (const constant of [
			'Private Const C = "3" Xor 1',
			'Private Const C = "1" & #1/1/2000#',
			'Private Const C = #12:00:00 PM# & 7',
			'Private Const C As Byte = 1 - "1"',
			'Private Const C = 922337203685477@ + 0.0001@',
			'Private Const C = "1" + "2"',
			'Private Const C As Integer = "2" + "3"',
		]) {
			expect(errors(constant), constant).toEqual([]);
		}
	});
});
