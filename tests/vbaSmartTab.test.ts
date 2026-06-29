import { describe, it, expect } from 'vitest';
import { smartTabShouldIndentLine } from '../src/vbaSmartTab';

const LINE = '    XlideAssert.IsFalse False';

describe('smartTabShouldIndentLine', () => {
	it('inserts a tab (does NOT indent the line) with the caret at end of content', () => {
		// The reported case: Tab at the end of a content line should not shift the line.
		expect(smartTabShouldIndentLine(LINE, LINE.length, true, false)).toBe(false);
	});

	it('inserts a tab with the caret in the middle of the content', () => {
		expect(smartTabShouldIndentLine(LINE, 12, true, false)).toBe(false);
	});

	it('indents the line when the caret is within the leading whitespace', () => {
		expect(smartTabShouldIndentLine(LINE, 0, true, false)).toBe(true);
		expect(smartTabShouldIndentLine(LINE, 4, true, false)).toBe(true); // at first non-ws
	});

	it('indents a blank or whitespace-only line', () => {
		expect(smartTabShouldIndentLine('', 0, true, false)).toBe(true);
		expect(smartTabShouldIndentLine('    ', 4, true, false)).toBe(true);
	});

	it('indents the block for a multi-line selection', () => {
		expect(smartTabShouldIndentLine(LINE, 5, false, true)).toBe(true);
	});

	it('does not indent for a within-line (single-line) selection', () => {
		expect(smartTabShouldIndentLine(LINE, 12, false, false)).toBe(false);
	});
});
