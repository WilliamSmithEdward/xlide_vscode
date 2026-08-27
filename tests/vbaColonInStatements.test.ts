import { describe, expect, it } from 'vitest';
import { analyzeVbaStructure } from '../src/vbaStructuralDiagnostics';
import { toLogicalLines } from '../src/vbaSourceScan';

// Issue #51, reported against 4.1.5: `If Range(Cell1:="a1") Is Nothing Then`
// reported "'End If' has no matching 'If'".
//
// A `:` separates statements in VBA, and the scanner split on every one of
// them. The colon in `:=` is not a separator, so the If header was torn into
// `If Range(Cell1` and `="a1") Is Nothing Then` - neither of which opens a
// block, leaving the matching `End If` unmatched.
//
// The same root cause reaches a second case the report did not name: the
// colons inside a date literal. `If Now > #12:30:00 PM# Then` produced the
// identical error.

const NL = '\n';

function structural(body: readonly string[]): string[] {
	return analyzeVbaStructure(['Sub T()', ...body, 'End Sub', ''].join(NL))
		.map((hit) => hit.message);
}

function segments(line: string): string[] {
	return toLogicalLines(line).logical.map((entry) => entry.text.trim());
}

describe('a colon that is not a statement separator (issue #51)', () => {
	it('accepts a named argument in an If condition', () => {
		expect(structural([
			'    If Range(Cell1:="a1") Is Nothing Then',
			'    End If',
		])).toEqual([]);
	});

	it('accepts it spaced, and more than one of them', () => {
		expect(structural([
			'    If Range(Cell1 := "a1") Is Nothing Then',
			'    End If',
		]), 'spaced').toEqual([]);
		expect(structural([
			'    If Range(Cell1:="a1", Cell2:="b2") Is Nothing Then',
			'    End If',
		]), 'two named arguments').toEqual([]);
	});

	it('accepts one in every other block header', () => {
		expect(structural([
			'    If Range("A1") Is Nothing Then',
			'    ElseIf Range(Cell1:="A2") Is Nothing Then',
			'    End If',
		]), 'ElseIf').toEqual([]);
		expect(structural([
			'    Do While Not Range(Cell1:="A1") Is Nothing',
			'    Loop',
		]), 'Do While').toEqual([]);
		expect(structural([
			'    Select Case Application.WorksheetFunction.Sum(Arg1:=1)',
			'    Case 1',
			'    End Select',
		]), 'Select Case').toEqual([]);
	});

	it('accepts a date or time literal, which the report did not name', () => {
		for (const literal of [
			'#12:30:00 PM#',
			'#12:30:00#',
			'#1/1/2020#',
			'#1/1/2020 12:30:00 PM#',
		]) {
			expect(structural([`    If Now > ${literal} Then`, '    End If']), literal).toEqual([]);
		}
	});
});

describe('the colons that ARE separators still are', () => {
	it('splits a real separator, including one after a named argument', () => {
		expect(segments('Application.Run Macro:="x": Debug.Print 1'))
			.toEqual(['Application.Run Macro:=', 'Debug.Print 1']);
	});

	it('leaves a file number alone but keeps the separator beside it', () => {
		// The date-literal match requires a leading digit and date punctuation
		// only, so `#1, "a": Close #` is not a literal and the colon between the
		// two statements survives. `#` opens a file number and a preprocessor
		// directive too, and neither may be swallowed.
		expect(segments('Print #1, "a": Close #1'))
			.toEqual(['Print #1,', 'Close #1']);
	});

	it('still splits a line label', () => {
		expect(structural(['    GoTo H', 'H:', '    Debug.Print 1'])).toEqual([]);
	});

	it('still treats a single-line If as one statement', () => {
		expect(structural(['    Dim x As Long', '    If x = 1 Then: x = 2'])).toEqual([]);
	});
});

describe('the structural findings this must not suppress', () => {
	it('still reports an End If with nothing open', () => {
		expect(structural(['    End If'])).toEqual([
			"'End If' has no matching 'If'.",
		]);
	});

	it('still reports an If left unclosed, named argument and all', () => {
		expect(structural(['    If Range(Cell1:="a1") Is Nothing Then'])).toEqual([
			"Missing 'End If' for 'If'.",
		]);
	});
});
