import { describe, expect, it } from 'vitest';
import { resolveDiagnosticCodeActions } from '../src/analyzer';
import { analyzeVbaModuleSource } from '../src/vbaModuleAnalysis';

// Issue #52, found from xlide_vbide: the suppression quick fix offered
// `disable-next-line` whatever the rule's scope, and for a MODULE-scoped rule
// the directive it wrote could not work.
//
// `option-explicit-missing` is evaluated for the whole module and anchors at
// (1,1). Inserting a line directive above line 1 re-anchors the finding onto
// the directive itself, which `disable-next-line` never covers - it covers the
// line BELOW. So the comment suppressed nothing and applying the fix again just
// stacked another directive.
//
// The action is now chosen from the rule's own `suppressionScopes`.

const NL = '\n';

function analyze(source: string) {
	return analyzeVbaModuleSource({
		source,
		moduleName: 'M',
		knownIdentifiers: new Set<string>(),
	} as never);
}

function applyEdits(
	source: string,
	edits: readonly { span: { start: number; end: number }; newText: string }[],
): string {
	let out = source;
	for (const edit of [...edits].sort((left, right) => right.span.start - left.span.start)) {
		out = out.slice(0, edit.span.start) + edit.newText + out.slice(edit.span.end);
	}
	return out;
}

/** The suppression fix offered for the first finding of `code`, applied. */
function suppress(source: string, code: string): { title: string; text: string } {
	const hit = analyze(source).diagnostics.find((entry) => entry.code === code);
	if (!hit) {
		throw new Error(`no ${code} finding to suppress`);
	}
	const action = resolveDiagnosticCodeActions(source, {
		code: hit.code,
		span: hit.span,
		includeSuppressionAction: true,
	} as never).find((fix) => fix.title.startsWith('Suppress'));
	if (!action) {
		throw new Error(`no suppression action for ${code}`);
	}
	return { title: action.title, text: applyEdits(source, action.edits) };
}

function counts(source: string, code: string): { live: number; suppressed: number } {
	const result = analyze(source);
	return {
		live: result.diagnostics.filter((entry) => entry.code === code).length,
		suppressed: result.suppressedDiagnostics.filter((entry) => entry.code === code).length,
	};
}

describe('a module-scoped rule gets the file directive', () => {
	const SOURCE = ['Public Sub P()', '    x = 1', 'End Sub', ''].join(NL);

	it('offers the file action, not the next-line one', () => {
		expect(suppress(SOURCE, 'option-explicit-missing').title)
			.toBe("Suppress 'option-explicit-missing' in this file");
	});

	it('writes a directive that actually suppresses the finding', () => {
		// The whole point. Before the fix this stayed live at 1.
		expect(counts(SOURCE, 'option-explicit-missing')).toEqual({ live: 1, suppressed: 0 });
		const { text } = suppress(SOURCE, 'option-explicit-missing');
		expect(text.split(NL)[0]).toBe("' @xlide-analysis-disable-file option-explicit-missing");
		expect(counts(text, 'option-explicit-missing')).toEqual({ live: 0, suppressed: 1 });
	});

	it('cannot stack, because the finding is gone', () => {
		const { text } = suppress(SOURCE, 'option-explicit-missing');
		expect(() => suppress(text, 'option-explicit-missing'))
			.toThrow(/no option-explicit-missing finding/);
	});

	it('lands after the Attribute header, where the directive is legal', () => {
		// `disable-file` must appear before the first non-comment, non-attribute
		// line; above the attributes would be a different kind of wrong.
		const withHeader = ['Attribute VB_Name = "M"', 'Public Sub P()', 'End Sub', ''].join(NL);
		const { text } = suppress(withHeader, 'option-explicit-missing');
		expect(text.split(NL).slice(0, 2)).toEqual([
			'Attribute VB_Name = "M"',
			"' @xlide-analysis-disable-file option-explicit-missing",
		]);
		expect(counts(text, 'option-explicit-missing')).toEqual({ live: 0, suppressed: 1 });
	});
});

describe('a positional rule keeps the next-line directive', () => {
	it('suppresses a finding on its own line', () => {
		const source = ['Option Explicit', 'Public Sub P()', '    nope = 1', 'End Sub', ''].join(NL);
		const { title, text } = suppress(source, 'undeclared-variable');
		expect(title).toBe("Suppress 'undeclared-variable' on next line");
		expect(text.split(NL)[2]).toBe("    ' @xlide-analysis-disable-next-line undeclared-variable");
		expect(counts(text, 'undeclared-variable')).toEqual({ live: 0, suppressed: 1 });
	});

	it('keeps the indentation of the line it guards', () => {
		const source = [
			'Option Explicit',
			'Public Sub P()',
			'    If True Then',
			'        nope = 1',
			'    End If',
			'End Sub',
			'',
		].join(NL);
		const { text } = suppress(source, 'undeclared-variable');
		expect(text.split(NL)[3]).toBe("        ' @xlide-analysis-disable-next-line undeclared-variable");
		expect(counts(text, 'undeclared-variable')).toEqual({ live: 0, suppressed: 1 });
	});
});
