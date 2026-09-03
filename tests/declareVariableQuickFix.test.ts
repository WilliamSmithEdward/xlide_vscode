// github.com/WilliamSmithEdward/xlide_vscode/issues/59.
//
// `undeclared-variable` is the most common finding in any Option Explicit
// project and had nothing on the lightbulb but the generic suppression, while
// the neighbouring `unknown-call` has offered to write the missing procedure
// for a long time.

import { describe, expect, it } from 'vitest';
import { resolveDiagnosticCodeActions, type VbaDiagnostic } from '../src/analyzer';
import { analyzeProjectModule } from './diagnostics/helpers';

function undeclared(lines: string[]): { source: string; diagnostics: VbaDiagnostic[] } {
	const source = lines.join('\r\n') + '\r\n';
	return {
		source,
		diagnostics: analyzeProjectModule(source, [{ moduleName: 'Mod1', source }], 'Mod1')
			.filter((d) => d.code === 'undeclared-variable'),
	};
}

/** The source after applying the declare fix, or undefined when none is offered. */
function applied(lines: string[]): string | undefined {
	const { source, diagnostics } = undeclared(lines);
	expect(diagnostics).toHaveLength(1);
	const fix = resolveDiagnosticCodeActions(source, diagnostics[0])
		.find((action) => action.title.startsWith('Declare'));
	if (!fix) {
		return undefined;
	}
	const edit = fix.edits[0];
	return source.slice(0, edit.span.start) + edit.newText + source.slice(edit.span.end);
}

describe('declaring the missing variable', () => {
	it('inserts the Dim under the declarations already there', () => {
		expect(applied([
			'Option Explicit', '', 'Public Sub Caller()',
			'    Dim total As Long', '    total = 1', '    missingName = 3', 'End Sub',
		])).toContain('    Dim total As Long\r\n    Dim missingName As Long\r\n');
	});

	it('inserts at the top of a procedure that has no declarations', () => {
		expect(applied([
			'Option Explicit', '', 'Public Sub Caller()', '    label = "hi"', 'End Sub',
		])).toContain('Public Sub Caller()\r\n    Dim label As String\r\n    label = "hi"');
	});

	it('takes the type from the assignment', () => {
		expect(applied(['Option Explicit', 'Sub T()', '    n = 3', 'End Sub']))
			.toContain('Dim n As Long');
		expect(applied(['Option Explicit', 'Sub T()', '    s = "a"', 'End Sub']))
			.toContain('Dim s As String');
		// `Set` says object whatever the right-hand side turns out to be.
		expect(applied(['Option Explicit', 'Sub T()', '    Set c = New Collection', 'End Sub']))
			.toContain('Dim c As Object');
	});

	it('falls back to Variant when nothing narrows the value', () => {
		expect(applied([
			'Option Explicit', 'Sub T()', '    Dim v As Variant', '    x = v', 'End Sub',
		])).toContain('Dim x As Variant');
	});

	// A bare read gives nothing to infer a type from, and declaring a name the
	// author only reads is as likely to paper over a typo as to be the fix.
	it('offers nothing for a name that is only read', () => {
		const { source, diagnostics } = undeclared([
			'Option Explicit', 'Sub T()', '    Debug.Print unknownThing', 'End Sub',
		]);
		expect(diagnostics).toHaveLength(1);
		expect(resolveDiagnosticCodeActions(source, diagnostics[0])
			.some((a) => a.title.startsWith('Declare'))).toBe(false);
	});

	it('keeps the suppression fix available alongside it', () => {
		const { source, diagnostics } = undeclared([
			'Option Explicit', 'Sub T()', '    n = 3', 'End Sub',
		]);
		const actions = resolveDiagnosticCodeActions(source, {
			...diagnostics[0],
			includeSuppressionAction: true,
		});
		expect(actions.some((a) => a.title.startsWith('Declare'))).toBe(true);
		expect(actions.some((a) => a.title.toLowerCase().includes('suppress'))).toBe(true);
	});
});
