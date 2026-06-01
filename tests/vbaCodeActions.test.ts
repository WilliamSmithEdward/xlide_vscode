import { describe, expect, it } from 'vitest';
import {
	analyzeModule,
	resolveDiagnosticCodeActions,
	type VbaDiagnostic,
	type VbaTextEdit,
} from '../src/analyzer';

function byCode(diags: readonly VbaDiagnostic[], code: string): VbaDiagnostic[] {
	return diags.filter((diag) => diag.code === code);
}

function firstDiagnostic(source: string, code: string): VbaDiagnostic {
	const diag = byCode(analyzeModule(source), code)[0];
	expect(diag).toBeTruthy();
	return diag;
}

function applyEdits(source: string, edits: readonly VbaTextEdit[]): string {
	return [...edits]
		.sort((a, b) => b.span.start - a.span.start)
		.reduce(
			(next, edit) =>
				next.slice(0, edit.span.start) + edit.newText + next.slice(edit.span.end),
			source,
		);
}

describe('resolveDiagnosticCodeActions', () => {
	it('adds parentheses around an explicit Call argument list', () => {
		const source = 'Sub T()\n    Call MsgBox "hello"\nEnd Sub\n';
		const diag = firstDiagnostic(source, 'call-requires-parens');

		const actions = resolveDiagnosticCodeActions(source, diag);

		expect(actions).toHaveLength(1);
		expect(actions[0].title).toBe('Add parentheses to Call argument list');
		expect(applyEdits(source, actions[0].edits)).toBe(
			'Sub T()\n    Call MsgBox("hello")\nEnd Sub\n',
		);
	});

	it('adds parentheses around an explicit member Call argument list before comments', () => {
		const source = "Sub T()\n    Call obj.Method 1, 2 ' keep comment\nEnd Sub\n";
		const diag = firstDiagnostic(source, 'call-requires-parens');

		const actions = resolveDiagnosticCodeActions(source, diag);

		expect(actions).toHaveLength(1);
		expect(applyEdits(source, actions[0].edits)).toBe(
			"Sub T()\n    Call obj.Method(1, 2) ' keep comment\nEnd Sub\n",
		);
	});

	it('removes empty parentheses from a standalone zero-argument runtime call', () => {
		const source = 'Sub T()\n    DoEvents()\nEnd Sub\n';
		const diag = firstDiagnostic(source, 'call-statement-forbids-parens');

		const actions = resolveDiagnosticCodeActions(source, diag);

		expect(actions).toHaveLength(1);
		expect(actions[0].title).toBe('Remove empty parentheses');
		expect(applyEdits(source, actions[0].edits)).toBe(
			'Sub T()\n    DoEvents\nEnd Sub\n',
		);
	});

	it('removes only the empty argument list from a standalone host member call', () => {
		const source = 'Sub T()\n    Application.Calculate()\nEnd Sub\n';
		const diag = firstDiagnostic(source, 'call-statement-forbids-parens');

		const actions = resolveDiagnosticCodeActions(source, diag);

		expect(actions).toHaveLength(1);
		expect(applyEdits(source, actions[0].edits)).toBe(
			'Sub T()\n    Application.Calculate\nEnd Sub\n',
		);
	});

	it('does not offer a call-parentheses fix unless the diagnostic span contains empty parentheses', () => {
		const source = 'Sub T()\n    DoEvents\nEnd Sub\n';
		const start = source.indexOf('DoEvents');
		const actions = resolveDiagnosticCodeActions(source, {
			code: 'call-statement-forbids-parens',
			span: { start, end: start + 'DoEvents'.length },
		});

		expect(actions).toHaveLength(0);
	});

	it('rewrites invalid explicit Call syntax for runtime statements that cannot use Call', () => {
		const source = 'Sub T()\n    Call DoEvents()\nEnd Sub\n';
		const diag = firstDiagnostic(source, 'invalid-explicit-call-target');

		const actions = resolveDiagnosticCodeActions(source, diag);

		expect(actions).toHaveLength(1);
		expect(actions[0].title).toBe('Use bare runtime call syntax');
		expect(applyEdits(source, actions[0].edits)).toBe(
			'Sub T()\n    DoEvents\nEnd Sub\n',
		);
	});

	it('removes only the Call keyword when invalid explicit Call syntax has no parentheses', () => {
		const source = "Sub T()\n    Call DoEvents ' keep pumping messages\nEnd Sub\n";
		const diag = firstDiagnostic(source, 'invalid-explicit-call-target');

		const actions = resolveDiagnosticCodeActions(source, diag);

		expect(actions).toHaveLength(1);
		expect(applyEdits(source, actions[0].edits)).toBe(
			"Sub T()\n    DoEvents ' keep pumping messages\nEnd Sub\n",
		);
	});

	it('adds Option Explicit at the top of a code module', () => {
		const source = 'Sub T()\nEnd Sub\n';
		const diag = firstDiagnostic(source, 'option-explicit-missing');

		const actions = resolveDiagnosticCodeActions(source, diag);

		expect(actions).toHaveLength(1);
		expect(actions[0].title).toBe('Add Option Explicit');
		expect(applyEdits(source, actions[0].edits)).toBe(
			'Option Explicit\nSub T()\nEnd Sub\n',
		);
	});

	it('adds Option Explicit after exported module attributes', () => {
		const source = 'Attribute VB_Name = "Module1"\n\nSub T()\nEnd Sub\n';
		const diag = firstDiagnostic(source, 'option-explicit-missing');

		const actions = resolveDiagnosticCodeActions(source, diag);

		expect(actions).toHaveLength(1);
		expect(applyEdits(source, actions[0].edits)).toBe(
			'Attribute VB_Name = "Module1"\nOption Explicit\n\nSub T()\nEnd Sub\n',
		);
	});

	it('does not add a duplicate Option Explicit for stale diagnostics', () => {
		const source = 'Option Explicit\nSub T()\nEnd Sub\n';
		const actions = resolveDiagnosticCodeActions(source, {
			code: 'option-explicit-missing',
			span: { start: 0, end: 0 },
		});

		expect(actions).toHaveLength(0);
	});
});
