// Diagnostics tests: edits on a sheet the code just protected (issue #471).
// Each sample was measured through pyVBAharness on 2026-10-02 in Excel 16.0
// (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode } from '../helpers/diagnostics';

const RANGE = 'host-argument-out-of-range';
const SETUP = ['Dim w2 As Worksheet', 'Set w2 = Worksheets.Add'];

function hits(...lines: string[]) {
	const src = `Option Explicit\nFunction Main() As Variant\n${[...SETUP, ...lines].map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
	return byCode(analyzeModule(src), RANGE);
}

describe('a sheet the code just protected (issue #471)', () => {
	it('does not infer cell locks, but retains an immediate wrong-password check', () => {
		for (const lines of [
			['w2.Protect', 'w2.Range("A1").Value = 1'],
			['w2.Protect', 'w2.Range("A1").Formula = "=1"'],
			['w2.Protect', 'w2.Range("A1").ClearContents'],
			['w2.Protect', 'w2.Rows(1).Insert'],
			['w2.Protect', 'w2.Range("A1").Font.Bold = True'],
			['w2.Protect "pw"', 'w2.Unprotect "nope"'],
		]) {
			const found = hits(...lines);
			expect(found, lines.join(' / ')).toHaveLength(lines[1].includes('Unprotect') ? 1 : 0);
		}
	});

	it('stays quiet on reads, the sheet itself, UserInterfaceOnly, Unprotect and unlocked cells', () => {
		for (const lines of [
			['w2.Protect', 'Main = w2.Range("A1").Value'],
			['w2.Protect', 'w2.Name = "Prot471"'],
			['w2.Protect UserInterfaceOnly:=True', 'w2.Range("A1").Value = 1'],
			['w2.Protect', 'w2.Unprotect', 'w2.Range("A1").Value = 1'],
			['w2.Protect "pw"', 'w2.Unprotect "pw"', 'w2.Range("A1").Value = 1'],
			['w2.Range("A1").Locked = False', 'w2.Protect', 'w2.Range("A1").Value = 1'],
		]) {
			expect(hits(...lines), lines.join(' / ')).toHaveLength(0);
		}
	});
});
