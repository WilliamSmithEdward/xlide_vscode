import { describe, expect, it, vi } from 'vitest';

vi.mock('vscode', async () => (await import('../helpers/vscodeMock')).vscodeMock());

import { shiftWithAutoDot, withAutoDotUnused, withSnippetAutoDot, type WithAutoDot } from '../../src/vbaWithAutoDot';
import { VBA_SMART_BLOCK_SNIPPETS, vbaSmartBlockSnippetsFor } from '../../src/vbaSmartBlockSnippets';
import { withMemberContinuationText } from '../../src/vbaSmartEnter';

// Issue #1620. The dot XLIDE puts on a With body line goes when the line is left
// with nothing after it; Excel rejects a lone `.` (oracle
// bare_leading_member_access_inside_with_compile).
const smartEnterDot: WithAutoDot = { line: 2, text: '        .', first: 1, last: 4, armed: true };

describe('the dot XLIDE puts on a With body line', () => {
	it('goes once the caret leaves its line, if nothing was typed after it', () => {
		expect(withAutoDotUnused(smartEnterDot, '        .', 2)).toBe(false);
		expect(withAutoDotUnused(smartEnterDot, '        .', 3)).toBe(true);
		expect(withAutoDotUnused(smartEnterDot, '        .', 0)).toBe(true);
		// A member typed after it, or the dot deleted, keeps the line as it is.
		expect(withAutoDotUnused(smartEnterDot, '        .Name = 1', 3)).toBe(false);
		expect(withAutoDotUnused(smartEnterDot, '        ', 3)).toBe(false);
	});

	it('waits, for a snippet, until the caret has been on it or has left the block', () => {
		const snippet = { ...smartEnterDot, armed: false };
		// Filling in the With's object on line 1.
		expect(withAutoDotUnused(snippet, '        .', 1)).toBe(false);
		expect(withAutoDotUnused(snippet, '        .', 4)).toBe(false);
		expect(withAutoDotUnused(snippet, '        .', 5)).toBe(true);
		expect(withAutoDotUnused(snippet, '        .', 0)).toBe(true);
	});

	it('moves with the edits above it and is dropped when its line is replaced', () => {
		// Two lines inserted above the block.
		expect(shiftWithAutoDot(smartEnterDot, { startLine: 0, endLine: 0, text: 'a\nb\n' })).toMatchObject({ line: 4, first: 3, last: 6 });
		// A line removed inside the block, below the dot.
		expect(shiftWithAutoDot(smartEnterDot, { startLine: 3, endLine: 4, text: '' })).toMatchObject({ line: 2, first: 1, last: 3 });
		// Enter at the end of the dot's line keeps it, and the block grows.
		expect(shiftWithAutoDot(smartEnterDot, { startLine: 2, endLine: 2, text: '\n        ' })).toMatchObject({ line: 2, last: 5 });
		// An edit below the block leaves it alone.
		expect(shiftWithAutoDot(smartEnterDot, { startLine: 9, endLine: 9, text: '\n' })).toEqual(smartEnterDot);
		expect(shiftWithAutoDot(smartEnterDot, { startLine: 1, endLine: 3, text: '' })).toBeUndefined();
	});

	it('is found in the With snippet in every block layout, and in no other insertion', () => {
		for (const layout of ['compact', 'comfy'] as const) {
			const spec = vbaSmartBlockSnippetsFor('statement', layout).find((candidate) => candidate.label === 'With');
			// The text the editor inserts once the snippet's placeholders are filled.
			const inserted = spec!.insertText.replace('${1:object}', 'ws').replace('$0', '');
			const dot = withSnippetAutoDot(inserted, 3);
			expect(dot, layout).toMatchObject({ text: '\t.', first: 3, armed: false });
			expect(inserted.split('\n')[dot!.line - 3], layout).toBe('\t.');
		}
		expect(VBA_SMART_BLOCK_SNIPPETS.some((candidate) => candidate.label === 'With')).toBe(true);
		const expanded = ['With ws', '\t.', 'End With'].join('\r\n');
		expect(withSnippetAutoDot(expanded, 7)).toEqual({ line: 8, text: '\t.', first: 7, last: 9, armed: false });
		const comfy = ['With ws', '', '\t.', '', 'End With'].join('\n');
		expect(withSnippetAutoDot(comfy, 0)).toMatchObject({ line: 2, first: 0, last: 4 });
		expect(withSnippetAutoDot(['If x Then', '\t.', 'End If'].join('\n'), 0)).toBeUndefined();
		expect(withSnippetAutoDot(['With ws', '\t.Name = 1', 'End With'].join('\n'), 0)).toBeUndefined();
		expect(withSnippetAutoDot('.', 0)).toBeUndefined();
	});

	it('seeds no second dot when Enter is pressed on a lone dot', () => {
		const lone = ['Sub T()', '    With ActiveSheet', '        .', '', '    End With', 'End Sub'].join('\n');
		expect(withMemberContinuationText(lone, 2)).toBeUndefined();
		const member = ['Sub T()', '    With ActiveSheet', '        .Range("A1")', '', '    End With', 'End Sub'].join('\n');
		expect(withMemberContinuationText(member, 2)).toBe('        .');
		const bracketed = ['Sub T()', '    With ActiveSheet', '        .[A1] = 1', '', '    End With', 'End Sub'].join('\n');
		expect(withMemberContinuationText(bracketed, 2)).toBe('        .');
	});
});
