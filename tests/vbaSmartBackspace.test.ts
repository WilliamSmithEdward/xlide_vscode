import { describe, expect, it } from 'vitest';
import { smartBackspaceShouldClearIndent } from '../src/vbaSmartBackspace';

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
