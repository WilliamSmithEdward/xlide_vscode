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
});
