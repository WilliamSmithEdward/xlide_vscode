// Which procedure a line belongs to, matching CodeModule.ProcOfLine
// (github.com/WilliamSmithEdward/xlide_vscode/issues/66). The awkward part is
// not the header: it is the comment and blank lines ABOVE a header, which the
// VBE gives to the procedure below them, not to the one above.

import { describe, expect, it } from 'vitest';
import {
	vbaProcedureLabelAtLine,
	vbaProcedureRanges,
	VBA_DECLARATIONS_LABEL,
} from '../src/vbaProcedureAtLine';

/** Every line's answer, so a rule cannot pass on the lines it was written for. */
function labels(source: string): string[] {
	return source.split('\n').map((_, line) => vbaProcedureLabelAtLine(source, line));
}

describe('the line a procedure starts owning', () => {
	const source = [
		/* 0 */ 'Option Explicit',
		/* 1 */ '',
		/* 2 */ "' Recalculates the sheet.",
		/* 3 */ 'Sub Recalculate()',
		/* 4 */ '    Debug.Print 1',
		/* 5 */ 'End Sub',
		/* 6 */ '',
		/* 7 */ "' Totals it.",
		/* 8 */ 'Private Function Total() As Long',
		/* 9 */ 'End Function',
		/* 10 */ '',
	].join('\n');

	it('gives a procedure the comment and blank lines above its header', () => {
		expect(labels(source)).toEqual([
			VBA_DECLARATIONS_LABEL,
			'Sub Recalculate',
			'Sub Recalculate',
			'Sub Recalculate',
			'Sub Recalculate',
			'Sub Recalculate',
			'Function Total',
			'Function Total',
			'Function Total',
			'Function Total',
			'Function Total',
		]);
	});

	it('gives the lines after the last End to that last procedure', () => {
		expect(vbaProcedureLabelAtLine(source, 10)).toBe('Function Total');
	});
});

describe('the declarations section', () => {
	it('runs to the last line of code above the first procedure lead-in', () => {
		const source = [
			'Option Explicit',
			'Public Total As Long',
			'',
			'Sub T()',
			'End Sub',
		].join('\n');
		expect(labels(source)).toEqual([
			VBA_DECLARATIONS_LABEL,
			VBA_DECLARATIONS_LABEL,
			'Sub T',
			'Sub T',
			'Sub T',
		]);
	});

	it('is the whole module when there is no procedure at all', () => {
		expect(labels('Option Explicit\nPublic X As Long'))
			.toEqual([VBA_DECLARATIONS_LABEL, VBA_DECLARATIONS_LABEL]);
	});

	it('holds a Declare, which is a declaration and not a procedure header', () => {
		const source = [
			'Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long',
			'Public Ticks As Long',
		].join('\n');
		expect(labels(source)).toEqual([VBA_DECLARATIONS_LABEL, VBA_DECLARATIONS_LABEL]);
	});
});

describe('the header shapes', () => {
	it('names each kind the way the source spells it', () => {
		const source = [
			'Property Get Name() As String',
			'End Property',
			'Property Let Name(v As String)',
			'End Property',
			'Property Set Thing(v As Object)',
			'End Property',
			'Public Static Function F()',
			'End Function',
			'Friend Sub S()',
			'End Sub',
		].join('\n');
		expect(vbaProcedureRanges(source).map((r) => `${r.kind} ${r.name}`)).toEqual([
			'Property Get Name',
			'Property Let Name',
			'Property Set Thing',
			'Function F',
			'Sub S',
		]);
	});

	it('reads a Sub written with no argument list', () => {
		expect(vbaProcedureLabelAtLine('Sub Bare\nEnd Sub', 0)).toBe('Sub Bare');
	});
});

describe('lines between two procedures that are neither comment nor blank', () => {
	it('leaves them with the procedure above, and stops the lead-in there', () => {
		const source = [
			/* 0 */ 'Sub A()',
			/* 1 */ 'End Sub',
			/* 2 */ 'Debug.Print 1',
			/* 3 */ '',
			/* 4 */ 'Sub B()',
			/* 5 */ 'End Sub',
		].join('\n');
		expect(labels(source)).toEqual(['Sub A', 'Sub A', 'Sub A', 'Sub B', 'Sub B', 'Sub B']);
	});

	it('never lets a lead-in run past the previous header', () => {
		// Two headers with only blanks between them: the second cannot swallow
		// the first, however inviting the blank line looks.
		const source = 'Sub A()\n\nSub B()\nEnd Sub';
		expect(labels(source)).toEqual(['Sub A', 'Sub B', 'Sub B', 'Sub B']);
	});
});

describe('line endings', () => {
	it('counts CRLF lines the same way it counts LF lines', () => {
		const lf = 'Option Explicit\n\nSub T()\nEnd Sub';
		const expected = [{ kind: 'Sub', name: 'T', firstLine: 1, lastLine: 3 }];
		expect(vbaProcedureRanges(lf)).toEqual(expected);
		expect(vbaProcedureRanges(lf.replace(/\n/g, '\r\n'))).toEqual(expected);
	});
});
