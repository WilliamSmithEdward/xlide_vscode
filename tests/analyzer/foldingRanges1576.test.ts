import { describe, expect, it } from 'vitest';
import { vbaFoldingRanges } from '../../src/analyzer/folding/foldingRanges';

/** One-based [first, last] lines, as an editor shows them. */
const folds = (lines: string[], text?: string[]): number[][] =>
	vbaFoldingRanges(lines.join('\r\n'), text?.join('\r\n')).map((range) => [range.start + 1, range.end + 1]);

describe('fold ranges (issue #1576)', () => {
	it('folds a block whole and each later arm up to the next, indented or flat', () => {
		const body = [
			['Sub M()', ''],
			['If a Then', '    '],
			['x = 1', '        '],
			['ElseIf b Then', '    '],
			['x = 2', '        '],
			['Else', '    '],
			['x = 3', '        '],
			['End If', '    '],
			['End Sub', ''],
		];
		const expected = [[1, 9], [2, 8], [4, 5], [6, 7]];
		expect(folds(body.map(([line]) => line))).toEqual(expected);
		expect(folds(body.map(([line, indent]) => indent + line))).toEqual(expected);
	});

	it('folds Select Case and its Case arms', () => {
		expect(folds(['Sub M()', 'Select Case n', 'Case 1', 'x = 1', 'Case 2, 3: x = 2', 'y = 2', 'Case Else', 'z = 3', 'End Select', 'End Sub']))
			.toEqual([[1, 10], [2, 9], [3, 4], [5, 6], [7, 8]]);
	});

	it('folds every kind of loop and With, and two blocks a Next closes on one line', () => {
		expect(folds(['Sub M()', 'For i = 1 To 2', 'For j = 1 To 2', 'x = 1', 'Next j, i', 'Do', 'x = 1', 'Loop', 'While a', 'x = 1', 'Wend', 'With o', '.x = 1', 'End With', 'End Sub']))
			.toEqual([[1, 15], [2, 5], [3, 5], [6, 8], [9, 11], [12, 14]]);
	});

	it('reads a block opened or closed after other statements, and gives a one-line block no fold', () => {
		expect(folds(['Sub M()', 'If a Then x = 1', 'For i = 1 To 2: x = 1: Next', 'L1: x = 1: For i = 1 To 2', 'y = 1', 'x = 1: Next', 'End Sub']))
			.toEqual([[1, 7], [4, 6]]);
	});

	it('folds #If to its #End If, and #ElseIf and #Else up to the next, at the margin of an indented procedure', () => {
		expect(folds(['Sub M()', '    x = 1', '#If VBA7 Then', '    y = 1', '#ElseIf Mac Then', '    y = 2', '#Else', '    y = 3', '#End If', 'End Sub']))
			.toEqual([[1, 10], [3, 9], [5, 6], [7, 8]]);
	});

	it('keeps the fold of a procedure whose header stands in #If arms, over the #If it crosses', () => {
		const ranges = folds(['#If VBA7 Then', 'Function F() As LongPtr', '#Else', 'Function F() As Long', '#End If', '    F = 1', 'End Function']);
		expect(ranges).toContainEqual([2, 7]);
		expect(ranges).not.toContainEqual([1, 5]);
	});

	it('folds Type and Enum, and keeps what only indentation folds', () => {
		expect(folds(['Type T', 'a As Long', 'End Type', 'Enum E', 'a', 'End Enum'])).toEqual([[1, 3], [4, 6]]);
		expect(folds(['Sub M()', '    x = a + _', '        b + _', '        c', '    y = 1', 'End Sub'])).toEqual([[1, 6], [2, 4]]);
	});

	it('gives a block not yet closed no fold of its own, and folds by indentation until it is', () => {
		expect(folds(['Sub M()', 'If a Then', 'x = 1'])).toEqual([]);
		expect(folds(['Sub M()', '    If a Then', '        x = 1', '    For i = 1 To 2', '        y = 2', 'x = 1'])).toEqual([[1, 5], [2, 3], [4, 5]]);
	});

	it('reads indentation from the text and blocks from the analysis source, as for a VB6 designer header', () => {
		const text = ['VERSION 5.00', 'Begin VB.Form F', '   Caption = "x"', 'End', 'Sub M()', 'x = 1', 'End Sub'];
		const blanked = ['', '', '', '', 'Sub M()', 'x = 1', 'End Sub'];
		expect(folds(blanked, text)).toEqual([[2, 3], [5, 7]]);
	});

	it('returns ranges nested and with one per first line, on the same answer each time', () => {
		const module = ['Sub M()', '    If a Then', '#If A Then', '        If b Then', '#Else', '        If c Then', '#End If', '            x = 1', '        End If', '    End If', 'End Sub'];
		const ranges = folds(module);
		expect(new Set(ranges.map(([start]) => start)).size).toBe(ranges.length);
		const stack: number[][] = [];
		for (const range of ranges) {
			while (stack.length > 0 && stack[stack.length - 1][1] < range[0]) { stack.pop(); }
			if (stack.length > 0) { expect(range[1]).toBeLessThanOrEqual(stack[stack.length - 1][1]); }
			stack.push(range);
		}
		expect(folds(module)).toEqual(ranges);
	});
});
