import { describe, expect, it } from 'vitest';
import {
	buildWorkbookAnalysisPlainText,
	buildWorkbookAnalysisResultsModel,
} from '../src/workbookAnalysisResultsModel';
import type { WorkbookAnalysisResult } from '../src/vbaWorkbookAnalysis';

function resultFixture(): WorkbookAnalysisResult {
	return {
		filePath: 'C:/work/Book.xlsm',
		moduleCount: 3,
		errorCount: 2,
		warningCount: 1,
		summary: {
			byCategory: {
				syntax: 1,
				semantic: 2,
			},
			byDiagnosticKind: {
				compile: 2,
				runtimeRisk: 1,
			},
			vbeCompileEquivalentCount: 2,
			nonVbeCompileEquivalentCount: 1,
			suppressedCount: 4,
		},
		problems: [
			{
				moduleName: 'Module1',
				moduleType: 'standard',
				line: 4,
				column: 2,
				endColumn: 8,
				severity: 'error',
				code: 'missing-block-closer',
				ruleTitle: 'Missing block closer',
				category: 'syntax',
				vbeCompileEquivalent: true,
				diagnosticKind: 'compile',
				specReference: 'MS-VBAL',
				message: 'Missing End Sub.',
			},
			{
				moduleName: 'Person',
				moduleType: 'class',
				line: 12,
				column: 5,
				endColumn: 9,
				severity: 'warning',
				code: 'missing-return-assignment',
				ruleTitle: 'Missing return assignment',
				category: 'semantic',
				vbeCompileEquivalent: false,
				diagnosticKind: 'runtimeRisk',
				message: 'Function does not assign its return value.',
			},
			{
				moduleName: 'Module1',
				moduleType: 'standard',
				line: 8,
				column: 3,
				endColumn: 13,
				severity: 'error',
				code: 'undeclared-variable',
				ruleTitle: 'Undeclared variable',
				category: 'semantic',
				vbeCompileEquivalent: true,
				diagnosticKind: 'compile',
				message: 'Variable is not declared.',
			},
		],
	};
}

describe('workbook analysis results model', () => {
	it('groups rows by module while preserving workbook analysis summary counts', () => {
		const model = buildWorkbookAnalysisResultsModel(resultFixture());

		expect(model).toMatchObject({
			workbookName: 'Book.xlsm',
			moduleCount: 3,
			totalProblems: 3,
			errorCount: 2,
			warningCount: 1,
			suppressedCount: 4,
			vbeCompileEquivalentCount: 2,
			nonVbeCompileEquivalentCount: 1,
		});
		expect(model.byCategory).toEqual([
			{ name: 'semantic', count: 2 },
			{ name: 'syntax', count: 1 },
		]);
		expect(model.byDiagnosticKind).toEqual([
			{ name: 'compile', count: 2 },
			{ name: 'runtimeRisk', count: 1 },
		]);
		expect(model.groups.map((group) => ({
			moduleName: group.moduleName,
			moduleIcon: group.moduleIcon,
			total: group.total,
			errors: group.errorCount,
			warnings: group.warningCount,
		}))).toEqual([
			{ moduleName: 'Module1', moduleIcon: 'M', total: 2, errors: 2, warnings: 0 },
			{ moduleName: 'Person', moduleIcon: 'C', total: 1, errors: 0, warnings: 1 },
		]);
	});

	it('sorts module groups in the same order as the XLIDE tree', () => {
		const fixture = resultFixture();
		const base = fixture.problems[0];
		fixture.problems = [
			{ ...base, moduleName: 'Person', moduleType: 'class' },
			{ ...base, moduleName: 'Module1', moduleType: 'standard' },
			{ ...base, moduleName: 'Sheet1', moduleType: 'document' },
			{ ...base, moduleName: 'UserForm1', moduleType: 'userform' },
		];

		const model = buildWorkbookAnalysisResultsModel(fixture);

		expect(model.groups.map((group) => `${group.moduleIcon}:${group.moduleName}`)).toEqual([
			'D:Sheet1',
			'F:UserForm1',
			'M:Module1',
			'C:Person',
		]);
	});

	it('builds a copyable report with module locations and rule codes', () => {
		const text = buildWorkbookAnalysisPlainText(buildWorkbookAnalysisResultsModel(resultFixture()));

		expect(text).toContain('XLIDE Analysis Results - Book.xlsm');
		expect(text).toContain('3 problem(s), 2 error(s), 1 warning(s), 3 module(s) checked.');
		expect(text).toContain('Suppressed by XLIDE analysis directives: 4.');
		expect(text).toContain('Module1:4:2 Missing End Sub. [missing-block-closer]');
		expect(text).toContain('Person:12:5 Function does not assign its return value. [missing-return-assignment]');
	});
});
