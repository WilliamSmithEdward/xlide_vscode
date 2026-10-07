import { describe, expect, it } from 'vitest';
import { emptyContinuedCommentMarkerStart, remainingBackspaceCleanup, smartBackspaceShouldClearIndent } from '../src/vbaSmartBackspace';

// Issue #43. XLIDE turns off `editor.trimAutoWhitespace` so a blank line keeps
// the indent the editor gave it, matching the VBE. That alone would make an
// unwanted blank line cost one Backspace per tab stop, so Backspace clears the
// whole indent at once - which is only observable two levels in, because at one
// level the editor's own tab stops already clear it in a single press.

describe('Backspace on a blank indented line', () => {
	it('clears the whole indent, at any depth', () => {
		expect(smartBackspaceShouldClearIndent('    ', 4, true)).toBe(true);
		expect(smartBackspaceShouldClearIndent('        ', 8, true)).toBe(true);
		expect(smartBackspaceShouldClearIndent(String.fromCharCode(9, 9), 2, true)).toBe(true);
	});

	it('leaves a line with content alone', () => {
		// Deleting the indent would move the text, which is not what Backspace is.
		expect(smartBackspaceShouldClearIndent('    Dim x', 9, true)).toBe(false);
		expect(smartBackspaceShouldClearIndent('    Dim x', 4, true)).toBe(false);
	});

	it('leaves column 1 and a wholly empty line alone', () => {
		expect(smartBackspaceShouldClearIndent('', 0, true)).toBe(false);
		expect(smartBackspaceShouldClearIndent('    ', 0, true)).toBe(false);
	});

	it('leaves a selection alone', () => {
		// A selected range is deleted as a range; the rule is for a bare caret.
		expect(smartBackspaceShouldClearIndent('        ', 8, false)).toBe(false);
	});
});


describe('cleanup following a native deletion', () => {
    it.each([
        ['            ', 12, undefined, 8, 12, '        ', { start: 0, end: 8 }],
        ['    \t  ', 5, undefined, 4, 5, '      ', { start: 0, end: 4 }],
        ["    ' ", 6, "    'note", 5, 6, "    '", { start: 4, end: 5 }],
        ["    '''' ", 9, "    ''''note", 8, 9, "    ''''", { start: 4, end: 8 }],
    ] as const)('clears only the remaining eligible prefix of %j', (before, caret, previous, start, end, after, expected) => {
        expect(remainingBackspaceCleanup(before, caret, previous, start, end, after)).toEqual(expected);
    });
    it.each([
        ['    a', 5, undefined, 4, 5, '    '], // The last ordinary character is not an indent deletion.
        ['    ', 4, undefined, 0, 4, ''], // Native tab-stop deletion already did everything.
        ['    ', 4, undefined, 2, 3, '   '], // Delete rather than Backspace at the saved caret.
        ['    ', 4, undefined, 3, 4, '  '], // A different edit must not be treated as this deletion.
        ['    ', 4, undefined, -1, 4, ''],
        ["    ' ", 6, 'value = 1', 5, 6, "    '"],
        ["    'content", 12, "    'note", 11, 12, "    'conten"],
    ] as const)('does no additional deletion for %j', (before, caret, previous, start, end, after) => {
        expect(remainingBackspaceCleanup(before, caret, previous, start, end, after)).toBeUndefined();
    });
    it('recognizes a continued marker without consuming its indentation', () => {
        expect(emptyContinuedCommentMarkerStart("    ''' ", 8, "    '''note")).toBe(4);
        expect(emptyContinuedCommentMarkerStart("    ''' ", 8, "    ''note")).toBeUndefined();
        expect(emptyContinuedCommentMarkerStart("    ' text", 6, "    'note")).toBeUndefined();
    });
});
