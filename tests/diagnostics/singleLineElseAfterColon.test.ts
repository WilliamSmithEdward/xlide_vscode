// Diagnostics tests: the Else of a one-line If after a colon,
// `If a Then b = 1: Else c = 2` (MS-VBAL 5.4.2.9 single-line-else-clause).
// The VBE verdicts are the if_single_line_else_after_colon_* cases in
// syntax_corpus/oracle/vbe_oracle_cases.json.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

import { byCode, expectDiagnostic, spanText } from '../helpers/diagnostics';
import { analyzeProjectModule } from './helpers';

describe('analyzeModule - else-without-if and a one-line If whose Else follows a colon', () => {
	const CODE = 'else-without-if';

	it('stays quiet for `If a Then b = 1: Else c = 2`', () => {
		const src = 'Sub T()\n    If a Then b = 1: Else c = 2\nEnd Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});

	it('stays quiet for `If a Then b = 1: Else: c = 2` (an empty statement after Else)', () => {
		const src = 'Sub T()\n    If a Then b = 1: Else: c = 2\nEnd Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});

	it('stays quiet when the colon continues onto the next line with `_`', () => {
		const src = 'Sub T()\n    If a Then b = 1: _\n        Else c = 2\nEnd Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});

	it('stays quiet when a block opened in the Then list closes before the Else', () => {
		// The parser splits `For`/`With` out of the tail into blocks of their own.
		const src =
			'Sub T()\n    If a Then For i = 1 To 2: Next: Else x = 1\n    If a Then With c: End With: Else d = 1\nEnd Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});

	it('stays quiet for a Select Case closed in the Then list before the Else', () => {
		const src = 'Sub T()\n    If a Then Select Case b: Case 1: c = 1: End Select: Else d = 1\nEnd Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});

	it('stays quiet for `Else If` written as tails: `If a Then b = 1: Else If x Then d = 2: Else e = 3`', () => {
		const src = 'Sub T()\n    If a Then b = 1: Else If x Then d = 2: Else e = 3\nEnd Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});

	it('stays quiet after a line label or line number', () => {
		const src = 'Sub T()\nL1: If a Then b = 1: Else c = 2\n10 If a Then b = 1: Else c = 2\nEnd Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});

	it('stays quiet in active code and ignores inactive `#If` code', () => {
		const src =
			'Sub T()\n#If True Then\n    If a Then b = 1: Else c = 2\n#Else\n    Else\n#End If\nEnd Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});

	it('reports no error at all under Option Explicit with every name declared', () => {
		const src =
			'Option Explicit\nSub T()\n    Dim a As Boolean, b As Long, c As Long\n    If a Then b = 1: Else c = 2\n    If a Then b = 1: Else: c = 2\nEnd Sub\n';
		const errors = analyzeProjectModule(src, [], 'Module1').filter((diag) => diag.severity === 'error');
		expect(errors).toEqual([]);
	});

	it('still reports an undeclared name in the Else tail', () => {
		// The tail is analysed as code, not skipped with the Else.
		const src =
			'Option Explicit\nSub T()\n    Dim a As Boolean, b As Long\n    If a Then b = 1: Else c = 2\nEnd Sub\n';
		const diagnostics = analyzeProjectModule(src, [], 'Module1');
		expect(byCode(diagnostics, CODE)).toHaveLength(0);
		expectDiagnostic(src, diagnostics, 'undeclared-variable', { span: 'c' });
	});

	it('stays quiet, and does not read the Else as a second clause, inside a block If', () => {
		const src =
			'Sub T()\n    If x Then\n        If a Then b = 1: Else c = 2\n    Else\n        y = 1\n    End If\nEnd Sub\n';
		const diagnostics = analyzeModule(src);
		expect(byCode(diagnostics, CODE)).toHaveLength(0);
		expect(byCode(diagnostics, 'else-branch-order')).toHaveLength(0);
	});

	it('still reports a second Else after a colon: `If a Then b = 1: Else c = 2: Else d = 3`', () => {
		const src = 'Sub T()\n    If a Then b = 1: Else c = 2: Else d = 3\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), CODE);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Else');
		expect(hits[0].span.start).toBe(src.indexOf('Else d'));
	});

	it('still reports an Else after a colon when the header has one: `If a Then b = 1 Else c = 2: Else d = 3`', () => {
		const src = 'Sub T()\n    If a Then b = 1 Else c = 2: Else d = 3\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), CODE);
		expect(hits).toHaveLength(1);
		expect(hits[0].span.start).toBe(src.indexOf('Else d'));
	});

	it('still reports a second Else after a colon inside a block If', () => {
		const src =
			'Sub T()\n    If x Then\n        If a Then b = 1: Else c = 2: Else d = 3\n    Else\n        y = 1\n    End If\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), CODE);
		expect(hits).toHaveLength(1);
		expect(hits[0].span.start).toBe(src.indexOf('Else d'));
	});

	it('still reports a stray Else on its own line', () => {
		const src = 'Sub T()\n    Else\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), CODE);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Else');
	});

	it('still reports an Else on the line after a one-line If (no continuation)', () => {
		const src = 'Sub T()\n    If a Then b = 1:\n    Else c = 2\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), CODE);
		expect(hits).toHaveLength(1);
		expect(hits[0].span.start).toBe(src.indexOf('Else'));
	});

	it('still reports a stray Else after a colon with no If on the line', () => {
		const src = 'Sub T()\n    b = 1: Else c = 2\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), CODE);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Else');
	});

	it('still reports `If a Then b = 1: ElseIf c Then d = 2`', () => {
		const src = 'Sub T()\n    If a Then b = 1: ElseIf c Then d = 2\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), CODE);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('ElseIf');
	});

	it('reports an ElseIf after a colon inside a block If too', () => {
		// The VBE says "Else without If"; main reported nothing here.
		const src = 'Sub T()\n    If x Then\n        If a Then b = 1: ElseIf c Then d = 2\n    End If\nEnd Sub\n';
		const diagnostics = analyzeModule(src);
		const hits = byCode(diagnostics, CODE);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('ElseIf');
		expect(byCode(diagnostics, 'else-branch-order')).toHaveLength(0);
	});

	it('lets the outer If take an Else after a colon when the header Else is a nested If\'s', () => {
		// `If b Then c = 1 Else d = 2` is the inner If with its Else; the outer
		// `If a` still has none, so `: Else e = 3` is its clause.
		const src = 'Sub T()\n    If a Then If b Then c = 1 Else d = 2: Else e = 3\nEnd Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});

	it('does the same inside a block If (main reports nothing there)', () => {
		const src =
			'Sub T()\n    If x Then\n        If a Then If b Then c = 1 Else d = 2: Else e = 3\n    End If\nEnd Sub\n';
		const diagnostics = analyzeModule(src);
		expect(byCode(diagnostics, CODE)).toHaveLength(0);
		expect(byCode(diagnostics, 'else-branch-order')).toHaveLength(0);
	});

	it('accepts two Else tails when two one-line Ifs are open', () => {
		const src = 'Sub T()\n    If a Then If b Then c = 1: Else d = 2: Else e = 3\nEnd Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});

	it('accepts an Else for a one-line If opened in a tail', () => {
		const src = 'Sub T()\n    If a Then b = 1: If c Then d = 2: Else e = 3\nEnd Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});

	it('accepts an Else for a one-line If opened in a tail after a header with its own Else, inside a block If', () => {
		const src =
			'Sub T()\n    If x Then\n        If a Then b Else c: If d Then e: Else f\n    End If\nEnd Sub\n';
		const diagnostics = analyzeModule(src);
		expect(byCode(diagnostics, CODE)).toHaveLength(0);
		expect(byCode(diagnostics, 'else-branch-order')).toHaveLength(0);
	});

	it('does not count End If as an opener: `If a Then b Else c: End If: Else d` still reported', () => {
		const src = 'Sub T()\n    If a Then b Else c: End If: Else d\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), CODE);
		expect(hits).toHaveLength(1);
		expect(hits[0].span.start).toBe(src.indexOf('Else d'));
	});

	it('does not count a member named If: the second Else is still reported', () => {
		const src = 'Sub T()\n    If a Then obj.If = 1: Else b: Else c\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), CODE);
		expect(hits).toHaveLength(1);
		expect(hits[0].span.start).toBe(src.indexOf('Else c'));
	});

	it('does not count a member named Else: the Else tail is accepted', () => {
		const src = 'Sub T()\n    If a Then obj.Else = 1: Else b\nEnd Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});

	it('recovers after a stray Else: a later one-line If on the line takes its own', () => {
		const src = 'Sub T()\n    If a Then b: Else c: Else d: If e Then f: Else g\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), CODE);
		expect(hits).toHaveLength(1);
		expect(hits[0].span.start).toBe(src.indexOf('Else d'));
	});

	it('leaves a block If with two Else clauses reported as before', () => {
		const src = 'Sub T()\n    If x Then\n        a = 1\n    Else\n        b = 2\n    Else\n        c = 3\n    End If\nEnd Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'else-branch-order', { span: 'Else' });
	});
});
