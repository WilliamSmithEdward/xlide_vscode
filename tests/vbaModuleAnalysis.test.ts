import { describe, expect, it } from 'vitest';
import { analyzeVbaModuleSource } from '../src/vbaModuleAnalysis';

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
