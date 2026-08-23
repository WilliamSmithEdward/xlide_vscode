import { describe, expect, it } from 'vitest';
import { analyzeModule } from '../src/analyzer';

// Issue #46, widened. A single-line `If` is ONE leaf statement, so every rule
// that reads a statement STRUCTURALLY - first token is the assignment target,
// first token is the callee - saw only `If` and never the statement being
// executed. Measured across the rule set: assignment-object-type-mismatch,
// const-assignment and argument-count were all blind, each a missed defect.
//
// The fix is opt-in per rule rather than folded into the shared walk, because a
// rule that SCANS statement text already sees inside a single-line If and would
// report the same defect twice (array-subscript-out-of-bounds does exactly that).

function codes(src: string, code: string): number {
	return analyzeModule(src, { host: 'excel' }).filter((d) => d.code === code).length;
}

const PRELUDE = [
	'Option Explicit',
	'Public Const K As Long = 1',
	'Public Sub T()',
	'    Dim ws As Worksheet',
	'    Dim r As Range',
	'    Dim x As Long',
	'    Set r = ActiveSheet.Range("A1")',
];
const EPILOGUE = [
	'End Sub',
	'Public Sub Helper(ByVal a As Long)',
	'End Sub',
	'Public Function Helper2(ByVal a As Long) As Long',
	'End Function',
	'',
];

function build(line: string): string {
	return [...PRELUDE, line, ...EPILOGUE].join('\n');
}

describe('rules read the statements a single-line If executes', () => {
	it.each([
		['assignment-object-type-mismatch', '    Set ws = r'],
		['const-assignment', '    K = 2'],
		['argument-count', '    Helper 1, 2, 3'],
		['argument-count', '    x = Helper2(1, 2, 3)'],
	])('%s reports the same defect on its own line and after Then', (code, line) => {
		const plain = codes(build(line), code);
		const inIf = codes(build(`    If True Then${line.trim() === line ? ' ' : ' '}${line.trim()}`), code);
		expect(plain, `${code} on its own line`).toBeGreaterThan(0);
		expect(inIf, `${code} after Then`).toBe(plain);
	});

	it('reports it once, not twice, for a rule that scans statement text', () => {
		// array-subscript-out-of-bounds reads the whole statement, so it always saw
		// inside the If; the branch pass must not make it report again.
		const src = [
			'Option Explicit',
			'Public Sub T()',
			'    Dim a(1 To 3) As Long',
			'    Dim i As Long',
			'    If i > 0 Then a(20) = 1',
			'End Sub',
			'',
		].join('\n');
		expect(codes(src, 'array-subscript-out-of-bounds')).toBe(1);
	});

	it('does not read an If CONDITION as an assignment', () => {
		// `If MAX = 10 Then` is a comparison. Visiting the condition as a statement
		// made it look like an assignment to a Const.
		const src = [
			'Option Explicit',
			'Public Const MAX As Long = 10',
			'Public Sub T()',
			'    If MAX = 10 Then Beep',
			'End Sub',
			'',
		].join('\n');
		expect(codes(src, 'const-assignment')).toBe(0);
	});
});
