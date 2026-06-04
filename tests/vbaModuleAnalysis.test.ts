import { describe, expect, it } from 'vitest';
import { analyzeVbaModuleSource } from '../src/vbaModuleAnalysis';
import { VBA_TEST_DIRECTIVE_DIAGNOSTIC_CODE } from '../src/vbaTestRunner';

function diagnosticsByCode(
	source: string,
	code: string,
	activeIncompleteExpressionOffset?: number,
) {
	return analyzeVbaModuleSource({
		source,
		activeIncompleteExpressionOffset,
	})
		.diagnostics
		.filter((diag) => diag.code === code);
}

describe('analyzeVbaModuleSource', () => {
	it('merges directive, structural, and semantic diagnostics through one suppression-aware core', () => {
		const source =
			'Option Explicit\n' +
			'Sub T()\n' +
			"    ' @xlide-analysis-disable-next-line undeclared-variable\n" +
			'    hiddenMissing = 1\n' +
			'    visibleMissing = 2\n' +
			'End Sub\n' +
			"' @xlide-analysis-disable-next-line not-a-rule\n" +
			'Sub Broken()\n';

		const result = analyzeVbaModuleSource({
			source,
			moduleName: 'Module1',
			knownIdentifiers: new Set<string>(),
		});

		expect(result.suppressedCount).toBe(1);
		expect(result.suppressedDiagnostics.map((diag) => diag.code)).toEqual(['undeclared-variable']);
		expect(result.diagnostics.map((diag) => diag.code)).toEqual([
			'analysis-suppression-directive',
			'missing-block-closer',
			'undeclared-variable',
		]);
		expect(result.diagnostics.find((diag) => diag.code === 'undeclared-variable')?.message)
			.toContain('visibleMissing');
	});

	it('honors project procedure context supplied by callers', () => {
		const source = 'Option Explicit\nSub T()\n    KnownProc 1\n    MissingProc 2\nEnd Sub\n';

		const result = analyzeVbaModuleSource({
			source,
			moduleName: 'Module1',
			knownProcedures: new Set(['knownproc']),
		});

		expect(result.diagnostics.map((diag) => diag.code)).toEqual(['unknown-call']);
		expect(result.diagnostics[0].message).toContain('MissingProc');
	});

	it('surfaces VBA test directive diagnostics through the shared analysis core', () => {
		const source =
			"' @xlide-test timeout=soon\n" +
			'Sub BadTimeout()\n' +
			'End Sub\n' +
			"' @xlide-test\n" +
			'Function NotRunnable() As Boolean\n' +
			'End Function\n';

		const result = analyzeVbaModuleSource({
			source,
			moduleName: 'Tests',
			moduleType: 'standard',
			severityOverrides: { 'option-explicit-missing': 'off' },
		});

		expect(result.diagnostics
			.filter((diag) => diag.code === VBA_TEST_DIRECTIVE_DIAGNOSTIC_CODE)
			.map((diag) => diag.message))
			.toEqual([
				'XLIDE test timeout must be a positive integer with optional ms or s suffix.',
				'XLIDE test directives must target a Sub procedure; Functions and Properties are not runnable tests.',
			]);
	});

	it('surfaces module declarations inside procedures through the shared analysis core', () => {
		const source =
			'Option Explicit\n' +
			'Sub T()\n' +
			'    Type LocalRecord\n' +
			'        Id As Long\n' +
			'    End Type\n' +
			'    Private Declare Function GetTickCount Lib "kernel32" () As Long\n' +
			'End Sub\n';

		const hits = diagnosticsByCode(source, 'module-declaration-in-procedure');

		expect(hits.map((hit) => source.slice(hit.span.start, hit.span.end))).toEqual([
			'Type',
			'Declare',
		]);
	});

	it('ignores inactive conditional procedure headers in structural analysis', () => {
		const source =
			'#If VBA7 Then\n' +
			'Public Function CreateFromPointer(ByVal p As LongPtr) As Object\n' +
			'#Else\n' +
			'Public Function CreateFromPointer(ByVal p As Long) As Object\n' +
			'#End If\n' +
			'    Set CreateFromPointer = Nothing\n' +
			'End Function\n';

		const result = analyzeVbaModuleSource({ source, moduleName: 'Callbacks' });

		expect(result.diagnostics.map((diag) => diag.code)).not.toContain('missing-block-closer');
		expect(result.diagnostics.map((diag) => diag.code)).not.toContain('unmatched-block-closer');
		expect(result.diagnostics.map((diag) => diag.code)).not.toContain('module-declaration-in-procedure');
	});

	it('ignores unknown-constant alternate procedure headers in structural analysis', () => {
		const source =
			'#If FULL_INTELLISENSE Then\n' +
			'Public Function AsAcc() As stdAcc\n' +
			'#Else\n' +
			'Public Function AsAcc() As Object\n' +
			'#End If\n' +
			'    Set AsAcc = Nothing\n' +
			'End Function\n';

		const result = analyzeVbaModuleSource({ source, moduleName: 'Window' });

		expect(result.diagnostics.map((diag) => diag.code)).not.toContain('missing-block-closer');
		expect(result.diagnostics.map((diag) => diag.code)).not.toContain('unmatched-block-closer');
		expect(result.diagnostics.map((diag) => diag.code)).not.toContain('module-declaration-in-procedure');
	});

	it('accepts colon-separated one-line procedures and block closers', () => {
		const source =
			'Public Function Run() As Variant: End Function\n' +
			'Public Sub Normalize()\n' +
			'    While InStr(1, value, "  ") > 0: value = Replace(value, "  ", " "): Wend\n' +
			'End Sub\n';

		const codes = analyzeVbaModuleSource({ source, moduleName: 'InlineBlocks' })
			.diagnostics.map((diag) => diag.code);

		expect(codes).not.toContain('missing-block-closer');
		expect(codes).not.toContain('unmatched-block-closer');
		expect(codes).not.toContain('module-declaration-in-procedure');
	});

	it('recovers after mismatched procedure closers without cascading into later declarations', () => {
		const source =
			'Public Property Get Name() As String\n' +
			'    Name = "Ada"\n' +
			'End Function\n' +
			'Public Sub Save()\n' +
			'End Sub\n';

		const result = analyzeVbaModuleSource({ source, moduleName: 'Person' });

		expect(result.diagnostics.map((diag) => diag.code)).toContain('missing-block-closer');
		expect(result.diagnostics.map((diag) => diag.code)).toContain('unmatched-block-closer');
		expect(result.diagnostics.map((diag) => diag.code)).not.toContain('module-declaration-in-procedure');
	});

	it('allows VBA test directive diagnostics to be suppressed explicitly', () => {
		const source =
			"' @xlide-analysis-disable-next-line vba-test-directive\n" +
			"' @xlide-test timeout=soon\n" +
			'Sub BadTimeout()\n' +
			'End Sub\n';

		const result = analyzeVbaModuleSource({
			source,
			moduleName: 'Tests',
			moduleType: 'standard',
			severityOverrides: { 'option-explicit-missing': 'off' },
		});

		expect(result.diagnostics.filter((diag) => diag.code === VBA_TEST_DIRECTIVE_DIAGNOSTIC_CODE))
			.toHaveLength(0);
		expect(result.suppressedDiagnostics.map((diag) => diag.code)).toEqual([VBA_TEST_DIRECTIVE_DIAGNOSTIC_CODE]);
	});

	it('suppresses only intentional deterministic runtime diagnostics for expected-error tests', () => {
		const source =
			'Option Explicit\n' +
			"' @xlide-test expected-error=13\n" +
			'Sub Test_ShouldRTE()\n' +
			'    Dim i As Integer\n' +
			'    i = "RTE"\n' +
			'    hiddenMissing = 1\n' +
			'End Sub\n' +
			'Sub Control()\n' +
			'    Dim j As Integer\n' +
			'    j = "RTE"\n' +
			'End Sub\n';

		const result = analyzeVbaModuleSource({
			source,
			moduleName: 'Tests',
			moduleType: 'standard',
			knownIdentifiers: new Set<string>(),
		});

		expect(result.suppressedDiagnostics.map((diag) => diag.code)).toEqual(['assignment-type-mismatch']);
		expect(result.diagnostics.filter((diag) => diag.code === 'assignment-type-mismatch'))
			.toHaveLength(1);
		expect(result.diagnostics.find((diag) => diag.code === 'assignment-type-mismatch')?.message)
			.toContain('j');
		expect(result.diagnostics.find((diag) => diag.code === 'undeclared-variable')?.message)
			.toContain('hiddenMissing');
	});

	it('supports expected-error any without hiding non-runtime diagnostics', () => {
		const source =
			'Option Explicit\n' +
			"' @xlide-test expected-error\n" +
			'Sub Test_ShouldRTE()\n' +
			'    Dim i As Integer\n' +
			'    i = "RTE"\n' +
			'    hiddenMissing = 1\n' +
			'End Sub\n';

		const result = analyzeVbaModuleSource({
			source,
			moduleName: 'Tests',
			moduleType: 'standard',
			knownIdentifiers: new Set<string>(),
		});

		expect(result.diagnostics.map((diag) => diag.code)).toEqual(['undeclared-variable']);
		expect(result.suppressedDiagnostics.map((diag) => diag.code)).toEqual(['assignment-type-mismatch']);
	});

	it('does not suppress deterministic runtime diagnostics with a mismatched expected error number', () => {
		const source =
			'Option Explicit\n' +
			"' @xlide-test expected-error=9\n" +
			'Sub Test_ShouldRTE()\n' +
			'    Dim i As Integer\n' +
			'    i = "RTE"\n' +
			'End Sub\n';

		const result = analyzeVbaModuleSource({
			source,
			moduleName: 'Tests',
			moduleType: 'standard',
		});

		expect(result.suppressedDiagnostics).toHaveLength(0);
		expect(result.diagnostics.map((diag) => diag.code)).toEqual(['assignment-type-mismatch']);
	});

	it('keeps hard diagnostics suppressed inside member ranges', () => {
		const source =
			'Option Explicit\n' +
			"' @xlide-analysis-disable-next-member undeclared-variable\n" +
			'Sub Hidden()\n' +
			'    hiddenMissing = 1\n' +
			'End Sub\n' +
			'Sub Visible()\n' +
			'    visibleMissing = 2\n' +
			'End Sub\n';

		const result = analyzeVbaModuleSource({
			source,
			moduleName: 'Module1',
			knownIdentifiers: new Set<string>(),
		});

		expect(result.suppressedCount).toBe(1);
		expect(result.suppressedDiagnostics).toHaveLength(1);
		expect(result.suppressedDiagnostics[0].message).toContain('hiddenMissing');
		expect(result.diagnostics).toHaveLength(1);
		expect(result.diagnostics[0].message).toContain('visibleMissing');
	});

	it('suppresses incomplete member access only for the active edit line', () => {
		const withSource =
			'Option Explicit\n' +
			'Sub T()\n' +
			'    With ActiveSheet\n' +
			'        .\n' +
			'    End With\n' +
			'End Sub\n';
		const withDotOffset = withSource.indexOf('        .') + '        .'.length;

		expect(diagnosticsByCode(withSource, 'invalid-expression-syntax')).toHaveLength(1);
		expect(diagnosticsByCode(
			withSource,
			'invalid-expression-syntax',
			withDotOffset,
		)).toHaveLength(0);

		const trailingSource = 'Option Explicit\nSub T()\n    ThisWorkbook.\nEnd Sub\n';
		const trailingOffset = trailingSource.indexOf('ThisWorkbook.') + 'ThisWorkbook.'.length;
		expect(diagnosticsByCode(trailingSource, 'invalid-expression-syntax')).toHaveLength(1);
		expect(diagnosticsByCode(
			trailingSource,
			'invalid-expression-syntax',
			trailingOffset,
		)).toHaveLength(0);

		const colonSource = 'Option Explicit\nSub T()\n    ThisWorkbook. : x = 1 *** 2\nEnd Sub\n';
		const colonDotOffset = colonSource.indexOf('ThisWorkbook.') + 'ThisWorkbook.'.length;
		const colonHits = analyzeVbaModuleSource({
			source: colonSource,
			activeIncompleteExpressionOffset: colonDotOffset,
		})
			.diagnostics
			.filter((diag) => diag.code === 'invalid-expression-syntax');
		expect(colonHits).toHaveLength(1);
		expect(colonHits[0].message).toContain('Invalid operator sequence');
	});

	it('suppresses trailing binary operators only for the active edit statement', () => {
		const source = 'Option Explicit\nSub T()\n    total = subtotal *\nEnd Sub\n';
		const operatorOffset = source.indexOf('*') + 1;
		const inactiveOffset = source.indexOf('End Sub');

		expect(diagnosticsByCode(source, 'invalid-expression-syntax')).toHaveLength(1);
		expect(diagnosticsByCode(
			source,
			'invalid-expression-syntax',
			operatorOffset,
		)).toHaveLength(0);
		expect(diagnosticsByCode(
			source,
			'invalid-expression-syntax',
			inactiveOffset,
		)).toHaveLength(1);
	});

	it('suppresses unmatched opening parentheses only for the active edit statement', () => {
		const source = 'Option Explicit\nSub T()\n    MsgBox(\nEnd Sub\n';
		const openParenOffset = source.indexOf('MsgBox(') + 'MsgBox('.length;
		const inactiveOffset = source.indexOf('End Sub');

		expect(diagnosticsByCode(source, 'unbalanced-parens')).toHaveLength(1);
		expect(diagnosticsByCode(
			source,
			'unbalanced-parens',
			openParenOffset,
		)).toHaveLength(0);
		expect(diagnosticsByCode(
			source,
			'unbalanced-parens',
			inactiveOffset,
		)).toHaveLength(1);
	});

	it('keeps active expression suppression local across colon-separated statements', () => {
		const source = 'Option Explicit\nSub T()\n    total = subtotal * : other = 1 *** 2\nEnd Sub\n';
		const firstOperatorOffset = source.indexOf('*') + 1;
		const hits = diagnosticsByCode(
			source,
			'invalid-expression-syntax',
			firstOperatorOffset,
		);

		expect(hits).toHaveLength(1);
		expect(hits[0].message).toContain('Invalid operator sequence');
	});
});
