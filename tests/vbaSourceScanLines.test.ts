// The line arithmetic every surface shares: the formatter, the quick fixes,
// the refactorings, project analysis and the test runner each used to carry
// a private copy of these.

import { describe, expect, it } from 'vitest';
import {
	lineIndexOf,
	lineStartAt,
	lineStartAtAnyBreak,
	lineStartOffsets,
	offsetToLineColumn,
	physicalLineSpanAtOffset,
	wholeLineSpan,
} from '../src/vbaSourceScan';

const SOURCE = 'Sub A()\r\n    x = 1\r\nEnd Sub';

describe('line helpers', () => {
	it('finds the line holding an offset', () => {
		const starts = lineStartOffsets(SOURCE);
		expect(starts).toEqual([0, 9, 20]);
		expect(lineIndexOf(starts, 0)).toBe(0);
		expect(lineIndexOf(starts, 8)).toBe(0);
		expect(lineIndexOf(starts, 9)).toBe(1);
		expect(lineIndexOf(starts, 19)).toBe(1);
		expect(lineIndexOf(starts, SOURCE.length)).toBe(2);
	});

	it('reports 1-based line and column', () => {
		const starts = lineStartOffsets(SOURCE);
		expect(offsetToLineColumn(starts, 0)).toEqual({ line: 1, column: 1 });
		expect(offsetToLineColumn(starts, 13)).toEqual({ line: 2, column: 5 });
		expect(offsetToLineColumn(starts, 20)).toEqual({ line: 3, column: 1 });
	});

	it('finds where a line starts', () => {
		expect(lineStartAt(SOURCE, 0)).toBe(0);
		expect(lineStartAt(SOURCE, 5)).toBe(0);
		expect(lineStartAt(SOURCE, 9)).toBe(9);
		expect(lineStartAt(SOURCE, 15)).toBe(9);
		expect(lineStartAt(SOURCE, SOURCE.length)).toBe(20);
	});

	it('answers 0 for offset 0 even when the text opens with a line break', () => {
		// A search from index 0 finds the leading LF; the line holding offset 0
		// still starts at 0.
		expect(lineStartAt('\nSub A()', 0)).toBe(0);
		expect(lineStartAtAnyBreak('\rSub A()', 0)).toBe(0);
		expect(physicalLineSpanAtOffset('\nSub A()', 0)).toEqual({ start: 0, end: 0 });
		expect(wholeLineSpan('\nSub A()', { start: 0, end: 0 })).toEqual({ start: 0, end: 1 });
	});

	it('breaks lines at a lone CR only when asked to', () => {
		const classic = 'a = 1\rb = 2';
		expect(lineStartAt(classic, 8)).toBe(0);
		expect(lineStartAtAnyBreak(classic, 8)).toBe(6);
		expect(lineStartAtAnyBreak(SOURCE, 15)).toBe(9);
	});

	it('widens a span to whole lines, the last break included', () => {
		expect(wholeLineSpan(SOURCE, { start: 13, end: 18 })).toEqual({ start: 9, end: 20 });
		expect(wholeLineSpan(SOURCE, { start: 22, end: 25 })).toEqual({ start: 20, end: SOURCE.length });
		expect(wholeLineSpan(SOURCE, { start: 2, end: 12 })).toEqual({ start: 0, end: 20 });
	});

	it('gives the physical line without its CR', () => {
		expect(physicalLineSpanAtOffset(SOURCE, 12)).toEqual({ start: 9, end: 18 });
		expect(SOURCE.slice(9, 18)).toBe('    x = 1');
		expect(physicalLineSpanAtOffset(SOURCE, 999)).toEqual({ start: 20, end: SOURCE.length });
	});
});
