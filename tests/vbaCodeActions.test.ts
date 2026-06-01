import { describe, expect, it } from 'vitest';
import {
	analyzeModule,
	ProjectIndex,
	resolveDiagnosticCodeActions,
	type AnalyzeModuleOptions,
	type VbaDiagnostic,
	type VbaTextEdit,
} from '../src/analyzer';
import { lintVbaSource, type VbaLintProblem } from '../src/vbaLinter';

function byCode(diags: readonly VbaDiagnostic[], code: string): VbaDiagnostic[] {
	return diags.filter((diag) => diag.code === code);
}

function firstDiagnostic(
	source: string,
	code: string,
	opts?: AnalyzeModuleOptions,
): VbaDiagnostic {
	const diag = byCode(analyzeModule(source, opts), code)[0];
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

function offsetAt(source: string, line: number, character: number): number {
	let currentLine = 0;
	for (let offset = 0; offset < source.length; offset++) {
		if (currentLine === line) {
			return offset + character;
		}
		if (source[offset] === '\n') {
			currentLine++;
		}
	}
	return source.length;
}

function lintAction(source: string, problem: VbaLintProblem) {
	return resolveDiagnosticCodeActions(source, {
		code: problem.code ?? '',
		message: problem.message,
		expectedClose: problem.expectedClose,
		insertLine: problem.insertLine,
		span: {
			start: offsetAt(source, problem.line, problem.startCol),
			end: offsetAt(source, problem.line, problem.endCol),
		},
	});
}

function projectClassMembers(
	modules: Array<{
		moduleName: string;
		source: string;
		moduleKind?: 'standard' | 'class' | 'document' | 'userform';
	}>,
): ReturnType<ProjectIndex['projectClassMembers']> {
	const project = new ProjectIndex();
	for (const mod of modules) {
		project.setModule({
			moduleName: mod.moduleName,
			moduleKind: mod.moduleKind ?? 'standard',
			source: mod.source,
		});
	}
	return project.projectClassMembers();
}

describe('resolveDiagnosticCodeActions', () => {
	it('inserts a missing procedure closer at EOF', () => {
		const source = 'Sub Foo()\n    MsgBox 1\n';
		const problem = lintVbaSource(source)[0];

		const actions = lintAction(source, problem);

		expect(actions).toHaveLength(1);
		expect(actions[0].title).toBe("Insert 'End Sub'");
		expect(applyEdits(source, actions[0].edits)).toBe(
			'Sub Foo()\n    MsgBox 1\nEnd Sub\n',
		);
	});

	it('inserts a missing procedure closer at EOF when the file has no final newline', () => {
		const source = 'Sub Foo()\n    MsgBox 1';
		const problem = lintVbaSource(source)[0];

		const actions = lintAction(source, problem);

		expect(actions).toHaveLength(1);
		expect(applyEdits(source, actions[0].edits)).toBe(
			'Sub Foo()\n    MsgBox 1\nEnd Sub\n',
		);
	});

	it('inserts an unclosed inner block before a mismatched outer closer', () => {
		const source = 'Sub Foo()\n    If x Then\n        y = 1\nEnd Sub\n';
		const problem = lintVbaSource(source).find(
			(candidate) => candidate.expectedClose === 'End If',
		);
		expect(problem).toBeTruthy();

		const actions = lintAction(source, problem!);

		expect(actions).toHaveLength(1);
		expect(actions[0].title).toBe("Insert 'End If'");
		expect(applyEdits(source, actions[0].edits)).toBe(
			'Sub Foo()\n    If x Then\n        y = 1\n    End If\nEnd Sub\n',
		);
	});

	it('does not offer structural insertion for unmatched closers', () => {
		const source = 'Sub Foo()\n    End If\nEnd Sub\n';
		const problem = lintVbaSource(source)[0];

		const actions = lintAction(source, problem);

		expect(problem.code).toBe('unmatched-block-closer');
		expect(actions).toHaveLength(0);
	});

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

	it('adds parentheses to a parenless function call used as an assignment expression', () => {
		const source =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n' +
			'Public Sub TestInvoiceTotal()\n' +
			'    total = InvoiceTotal 100, 0.08\n' +
			'End Sub\n';
		const diag = firstDiagnostic(source, 'expression-call-requires-parens');

		const actions = resolveDiagnosticCodeActions(source, diag);

		expect(actions).toHaveLength(1);
		expect(actions[0].title).toBe('Add parentheses to function call arguments');
		expect(applyEdits(source, actions[0].edits)).toContain(
			'    total = InvoiceTotal(100, 0.08)\n',
		);
	});

	it('does not wrap a parenless expression call when the argument boundary is ambiguous', () => {
		const source = 'Sub T()\n    total = 1 + InvoiceTotal 100, 0.08\nEnd Sub\n';
		const start = source.indexOf('InvoiceTotal');
		const actions = resolveDiagnosticCodeActions(source, {
			code: 'expression-call-requires-parens',
			span: { start, end: start + 'InvoiceTotal'.length },
		});

		expect(actions).toHaveLength(0);
	});

	it('adds Set to a proven object assignment', () => {
		const source =
			'Public Sub T()\n' +
			'    Dim ws As Worksheet\n' +
			'    ws = ActiveSheet\n' +
			'End Sub\n';
		const diag = firstDiagnostic(source, 'set-required');

		const actions = resolveDiagnosticCodeActions(source, diag);

		expect(actions).toHaveLength(1);
		expect(actions[0].title).toBe('Add Set to object assignment');
		expect(applyEdits(source, actions[0].edits)).toBe(
			'Public Sub T()\n' +
			'    Dim ws As Worksheet\n' +
			'    Set ws = ActiveSheet\n' +
			'End Sub\n',
		);
	});

	it('replaces Let with Set for a proven object assignment', () => {
		const source =
			'Public Sub T()\n' +
			'    Dim ws As Worksheet\n' +
			'    Let ws = ActiveSheet\n' +
			'End Sub\n';
		const diag = firstDiagnostic(source, 'set-required');

		const actions = resolveDiagnosticCodeActions(source, diag);

		expect(actions).toHaveLength(1);
		expect(actions[0].title).toBe('Replace Let with Set');
		expect(applyEdits(source, actions[0].edits)).toBe(
			'Public Sub T()\n' +
			'    Dim ws As Worksheet\n' +
			'    Set ws = ActiveSheet\n' +
			'End Sub\n',
		);
	});

	it('removes Set from a proven scalar assignment', () => {
		const source =
			'Public Sub T()\n' +
			'    Dim count As Integer\n' +
			'    Set count = 2\n' +
			'End Sub\n';
		const diag = firstDiagnostic(source, 'set-requires-object');

		const actions = resolveDiagnosticCodeActions(source, diag);

		expect(actions).toHaveLength(1);
		expect(actions[0].title).toBe('Remove Set from scalar assignment');
		expect(applyEdits(source, actions[0].edits)).toBe(
			'Public Sub T()\n' +
			'    Dim count As Integer\n' +
			'    count = 2\n' +
			'End Sub\n',
		);
	});

	it('adds Set to a proven object-valued member assignment', () => {
		const source =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    p.Child = New Person\n' +
			'End Sub\n';
		const diag = firstDiagnostic(source, 'set-required', {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'Person', moduleKind: 'class', source: 'Public Child As Person\n' },
			]),
		});

		const actions = resolveDiagnosticCodeActions(source, diag);

		expect(actions).toHaveLength(1);
		expect(applyEdits(source, actions[0].edits)).toBe(
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    Set p.Child = New Person\n' +
			'End Sub\n',
		);
	});

	it('removes Set from a proven scalar member assignment', () => {
		const source =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    Set p.Age = 2\n' +
			'End Sub\n';
		const diag = firstDiagnostic(source, 'set-requires-object', {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'Person', moduleKind: 'class', source: 'Public Age As Integer\n' },
			]),
		});

		const actions = resolveDiagnosticCodeActions(source, diag);

		expect(actions).toHaveLength(1);
		expect(applyEdits(source, actions[0].edits)).toBe(
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    p.Age = 2\n' +
			'End Sub\n',
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
