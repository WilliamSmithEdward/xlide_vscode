import { describe, expect, it } from 'vitest';
import {
	analyzeModule,
	filterDiagnosticsWithSuppressions,
	LINT_SUPPRESSION_DIRECTIVE_CODE,
	scanLintSuppressions,
	type VbaDiagnostic,
} from '../src/analyzer';
import { lintVbaSource } from '../src/vbaLinter';

function semanticDiagnostics(source: string): VbaDiagnostic[] {
	return analyzeModule(source, {
		knownIdentifiers: new Set<string>(),
	});
}

function diagnosticsByCode(diags: readonly VbaDiagnostic[], code: string): VbaDiagnostic[] {
	return diags.filter((diag) => diag.code === code);
}

describe('XLIDE lint suppression directives', () => {
	it('suppresses only the next physical line for a matching diagnostic code', () => {
		const source =
			'Option Explicit\n' +
			'Sub T()\n' +
			"    ' @xlide-lint-disable-next-line undeclared-variable\n" +
			'    notDeclared = 1\n' +
			'    stillMissing = 2\n' +
			'End Sub\n';

		const result = filterDiagnosticsWithSuppressions(source, semanticDiagnostics(source));

		expect(result.suppressedCount).toBe(1);
		expect(diagnosticsByCode(result.diagnostics, 'undeclared-variable')).toHaveLength(1);
		expect(result.directiveDiagnostics).toHaveLength(0);
	});

	it('suppresses diagnostics on the same physical line with disable-line', () => {
		const source =
			'Option Explicit\n' +
			'Sub T()\n' +
			"    notDeclared = 1 ' @xlide-lint-disable-line undeclared-variable\n" +
			'End Sub\n';

		const result = filterDiagnosticsWithSuppressions(source, semanticDiagnostics(source));

		expect(result.suppressedCount).toBe(1);
		expect(diagnosticsByCode(result.diagnostics, 'undeclared-variable')).toHaveLength(0);
	});

	it('treats an omitted diagnostic-code list as all diagnostics', () => {
		const source =
			'Option Explicit\n' +
			'Sub T()\n' +
			"    ' @xlide-lint-disable-next-line\n" +
			'    notDeclared = 1\n' +
			'End Sub\n';

		const result = filterDiagnosticsWithSuppressions(source, semanticDiagnostics(source));

		expect(result.suppressedCount).toBe(1);
		expect(result.diagnostics).toHaveLength(0);
	});

	it('supports file-level suppression before the first source line', () => {
		const source =
			'Attribute VB_Name = "Module1"\n' +
			" ' @xlide-lint-disable-file undeclared-variable -- generated shim\n" +
			'Option Explicit\n' +
			'Sub T()\n' +
			'    notDeclared = 1\n' +
			'End Sub\n';

		const result = filterDiagnosticsWithSuppressions(source, semanticDiagnostics(source));

		expect(result.suppressedCount).toBe(1);
		expect(result.diagnostics).toHaveLength(0);
		expect(result.directiveDiagnostics).toHaveLength(0);
	});

	it('reports and ignores late file-level suppression directives', () => {
		const source =
			'Option Explicit\n' +
			" ' @xlide-lint-disable-file all\n" +
			'Sub T()\n' +
			'    notDeclared = 1\n' +
			'End Sub\n';

		const result = filterDiagnosticsWithSuppressions(source, semanticDiagnostics(source));

		expect(result.suppressedCount).toBe(0);
		expect(diagnosticsByCode(result.diagnostics, 'undeclared-variable')).toHaveLength(1);
		expect(result.directiveDiagnostics).toHaveLength(1);
		expect(result.directiveDiagnostics[0].code).toBe(LINT_SUPPRESSION_DIRECTIVE_CODE);
	});

	it('ignores triple-apostrophe documentation comments', () => {
		const source =
			"''' @xlide-lint-disable-next-line missing-block-closer\n" +
			'Sub T()\n';
		const suppressions = scanLintSuppressions(source);
		const visible = lintVbaSource(source).filter((problem) =>
			!suppressions.isSuppressed(problem.code, problem.line),
		);

		expect(visible.map((problem) => problem.code)).toContain('missing-block-closer');
		expect(suppressions.diagnostics).toHaveLength(0);
	});

	it('ignores Rem comments', () => {
		const source =
			'Rem @xlide-lint-disable-next-line missing-block-closer\n' +
			'Sub T()\n';
		const suppressions = scanLintSuppressions(source);
		const visible = lintVbaSource(source).filter((problem) =>
			!suppressions.isSuppressed(problem.code, problem.line),
		);

		expect(visible.map((problem) => problem.code)).toContain('missing-block-closer');
		expect(suppressions.diagnostics).toHaveLength(0);
	});

	it('suppresses structural diagnostics through the same physical-line predicate', () => {
		const source =
			"' @xlide-lint-disable-next-line missing-block-closer\n" +
			'Sub T()\n';
		const suppressions = scanLintSuppressions(source);
		let suppressedCount = 0;
		const visible = lintVbaSource(source).filter((problem) => {
			if (suppressions.isSuppressed(problem.code, problem.line)) {
				suppressedCount++;
				return false;
			}
			return true;
		});

		expect(suppressedCount).toBe(1);
		expect(visible).toHaveLength(0);
	});

	it('reports unknown diagnostic codes and does not suppress with them', () => {
		const source =
			'Option Explicit\n' +
			'Sub T()\n' +
			"    ' @xlide-lint-disable-next-line not-a-rule\n" +
			'    notDeclared = 1\n' +
			'End Sub\n';

		const result = filterDiagnosticsWithSuppressions(source, semanticDiagnostics(source));

		expect(result.suppressedCount).toBe(0);
		expect(diagnosticsByCode(result.diagnostics, 'undeclared-variable')).toHaveLength(1);
		expect(result.directiveDiagnostics).toHaveLength(1);
		expect(result.directiveDiagnostics[0].message).toContain('not-a-rule');
	});

	it('keeps known codes active when a mixed list also contains an unknown code', () => {
		const source =
			'Option Explicit\n' +
			'Sub T()\n' +
			"    ' @xlide-lint-disable-next-line undeclared-variable, not-a-rule\n" +
			'    notDeclared = 1\n' +
			'End Sub\n';

		const result = filterDiagnosticsWithSuppressions(source, semanticDiagnostics(source));

		expect(result.suppressedCount).toBe(1);
		expect(result.diagnostics).toHaveLength(0);
		expect(result.directiveDiagnostics).toHaveLength(1);
	});

	it('reports malformed all-plus-code lists and does not guess', () => {
		const source =
			'Option Explicit\n' +
			'Sub T()\n' +
			"    ' @xlide-lint-disable-next-line all, undeclared-variable\n" +
			'    notDeclared = 1\n' +
			'End Sub\n';

		const result = filterDiagnosticsWithSuppressions(source, semanticDiagnostics(source));

		expect(result.suppressedCount).toBe(0);
		expect(diagnosticsByCode(result.diagnostics, 'undeclared-variable')).toHaveLength(1);
		expect(result.directiveDiagnostics).toHaveLength(1);
	});

	it('suppresses diagnostics inside the next supported top-level member only', () => {
		const source =
			'Option Explicit\n' +
			"' @xlide-lint-disable-next-member undeclared-variable\n" +
			'Sub Suppressed()\n' +
			'    memberMissing = 1\n' +
			'End Sub\n' +
			'Sub Visible()\n' +
			'    visibleMissing = 2\n' +
			'End Sub\n';

		const result = filterDiagnosticsWithSuppressions(source, semanticDiagnostics(source));

		expect(result.suppressedCount).toBe(1);
		expect(diagnosticsByCode(result.diagnostics, 'undeclared-variable')).toHaveLength(1);
		expect(result.diagnostics[0].message).toContain('visibleMissing');
		expect(result.directiveDiagnostics).toHaveLength(0);
	});

	it('reports disable-next-member when it is not directly before a supported member', () => {
		const source =
			'Option Explicit\n' +
			"' @xlide-lint-disable-next-member all\n" +
			'Dim moduleValue As Long\n' +
			'Sub T()\n' +
			'    notDeclared = 1\n' +
			'End Sub\n';

		const result = filterDiagnosticsWithSuppressions(source, semanticDiagnostics(source));

		expect(result.suppressedCount).toBe(0);
		expect(diagnosticsByCode(result.diagnostics, 'undeclared-variable')).toHaveLength(1);
		expect(result.directiveDiagnostics).toHaveLength(1);
		expect(result.directiveDiagnostics[0].message).toContain('disable-next-member');
	});

	it('supports Type and Enum members through the same next-member range model', () => {
		const source =
			"' @xlide-lint-disable-next-member invalid-declaration-name\n" +
			'Type TPoint\n' +
			'    X As Long\n' +
			'End Type\n' +
			"' @xlide-lint-disable-next-member invalid-declaration-name\n" +
			'Enum Color\n' +
			'    Red\n' +
			'End Enum\n' +
			'Sub T()\n' +
			'End Sub\n';
		const diagnostics = [
			{
				code: 'invalid-declaration-name',
				span: {
					start: source.indexOf('X As Long'),
					end: source.indexOf('X As Long') + 1,
				},
			},
			{
				code: 'invalid-declaration-name',
				span: {
					start: source.indexOf('Red'),
					end: source.indexOf('Red') + 3,
				},
			},
		];

		const result = filterDiagnosticsWithSuppressions(source, diagnostics);

		expect(result.suppressedCount).toBe(2);
		expect(result.diagnostics).toHaveLength(0);
		expect(result.directiveDiagnostics).toHaveLength(0);
	});

	it('suppresses diagnostics between matching block directives only', () => {
		const source =
			'Option Explicit\n' +
			'Sub T()\n' +
			"    ' @xlide-lint-disable-block undeclared-variable\n" +
			'    blockMissing = 1\n' +
			"    ' @xlide-lint-enable-block undeclared-variable\n" +
			'    afterMissing = 2\n' +
			'End Sub\n';

		const result = filterDiagnosticsWithSuppressions(source, semanticDiagnostics(source));

		expect(result.suppressedCount).toBe(1);
		expect(diagnosticsByCode(result.diagnostics, 'undeclared-variable')).toHaveLength(1);
		expect(result.diagnostics[0].message).toContain('afterMissing');
		expect(result.directiveDiagnostics).toHaveLength(0);
	});

	it('supports nested block directives with an explicit stack contract', () => {
		const source =
			'Sub T()\n' +
			"    ' @xlide-lint-disable-block undeclared-variable\n" +
			'    firstMissing = 1\n' +
			"    ' @xlide-lint-disable-block invalid-expression-syntax\n" +
			'    total = 1 *** 2\n' +
			"    ' @xlide-lint-enable-block invalid-expression-syntax\n" +
			'    secondMissing = 2\n' +
			"    ' @xlide-lint-enable-block undeclared-variable\n" +
			'    outsideMissing = 3\n' +
			'End Sub\n';
		const diagnostics = [
			{
				code: 'undeclared-variable',
				span: {
					start: source.indexOf('firstMissing'),
					end: source.indexOf('firstMissing') + 'firstMissing'.length,
				},
			},
			{
				code: 'invalid-expression-syntax',
				span: {
					start: source.indexOf('***'),
					end: source.indexOf('***') + 3,
				},
			},
			{
				code: 'undeclared-variable',
				span: {
					start: source.indexOf('secondMissing'),
					end: source.indexOf('secondMissing') + 'secondMissing'.length,
				},
			},
			{
				code: 'undeclared-variable',
				span: {
					start: source.indexOf('outsideMissing'),
					end: source.indexOf('outsideMissing') + 'outsideMissing'.length,
				},
			},
		];

		const result = filterDiagnosticsWithSuppressions(source, diagnostics);

		expect(result.suppressedCount).toBe(3);
		expect(result.diagnostics).toHaveLength(1);
		expect(result.diagnostics[0].span.start).toBe(source.indexOf('outsideMissing'));
		expect(result.directiveDiagnostics).toHaveLength(0);
	});

	it('reports stray enable-block directives and does not suppress', () => {
		const source =
			'Option Explicit\n' +
			'Sub T()\n' +
			"    ' @xlide-lint-enable-block undeclared-variable\n" +
			'    notDeclared = 1\n' +
			'End Sub\n';

		const result = filterDiagnosticsWithSuppressions(source, semanticDiagnostics(source));

		expect(result.suppressedCount).toBe(0);
		expect(diagnosticsByCode(result.diagnostics, 'undeclared-variable')).toHaveLength(1);
		expect(result.directiveDiagnostics).toHaveLength(1);
		expect(result.directiveDiagnostics[0].message).toContain('no matching');
	});

	it('reports unbalanced disable-block directives and does not suppress the open-ended block', () => {
		const source =
			'Option Explicit\n' +
			'Sub T()\n' +
			"    ' @xlide-lint-disable-block undeclared-variable\n" +
			'    notDeclared = 1\n' +
			'End Sub\n';

		const result = filterDiagnosticsWithSuppressions(source, semanticDiagnostics(source));

		expect(result.suppressedCount).toBe(0);
		expect(diagnosticsByCode(result.diagnostics, 'undeclared-variable')).toHaveLength(1);
		expect(result.directiveDiagnostics).toHaveLength(1);
		expect(result.directiveDiagnostics[0].message).toContain('missing a matching');
	});

	it('reports mismatched block close lists without guessing which block closed', () => {
		const source =
			'Option Explicit\n' +
			'Sub T()\n' +
			"    ' @xlide-lint-disable-block undeclared-variable\n" +
			'    notDeclared = 1\n' +
			"    ' @xlide-lint-enable-block argument-count\n" +
			'End Sub\n';

		const result = filterDiagnosticsWithSuppressions(source, semanticDiagnostics(source));

		expect(result.suppressedCount).toBe(0);
		expect(diagnosticsByCode(result.diagnostics, 'undeclared-variable')).toHaveLength(1);
		expect(result.directiveDiagnostics).toHaveLength(2);
		expect(result.directiveDiagnostics.map((diag) => diag.message).join('\n')).toContain('must match');
	});
});
