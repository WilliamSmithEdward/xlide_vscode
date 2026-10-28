// Issue #1494, run through the VBE oracle in Excel, Word and Access 16.0
// (build 20527): a document module named bare is its module's object. Into a
// variable or a parameter of another host class, a class of the project or
// an interface it does not implement, it raises 13; into its own class,
// Object, Variant or an interface it implements, it runs. A local of the
// same name shadows it. The cases are the issue1494_* fixtures in
// syntax_corpus/oracle/vbe_oracle_cases.json. PowerPoint has no document
// modules.

import * as fs from 'node:fs';
import * as path from 'node:path';
import { describe, expect, it } from 'vitest';
import { getAccessObjectModel } from '../../../src/analyzer/host/accessObjectModel';
import { getExcelObjectModel } from '../../../src/analyzer/host/excelObjectModel';
import { extractVariable } from '../../../src/analyzer/refactor/extractVariable';
import { applyVbaTextEdits } from '../../../src/analyzer/refactor/refactorTypes';
import { analyzeVbaModuleSource } from '../../../src/vbaModuleAnalysis';
import { buildVbaProjectIndex, effectiveModuleKind, projectAnalysisOptionsForModule, projectProcedureSignatures } from '../../../src/vbaProjectAnalysis';
import { oracleModuleInput, type OracleFixtureModule } from '../../helpers/oracleModuleInput';

interface OracleCase { id: string; expected: 'accepted' | 'rejected'; host?: string; modules: OracleFixtureModule[] }

const corpus: { cases: OracleCase[] } = JSON.parse(
	fs.readFileSync(path.join(process.cwd(), 'syntax_corpus', 'oracle', 'vbe_oracle_cases.json'), 'utf8'),
);
const measured = corpus.cases.filter((oracleCase) => oracleCase.id.startsWith('issue1494_'));

function errors(oracleCase: OracleCase): string[] {
	const modules = oracleCase.modules.map(oracleModuleInput);
	const project = buildVbaProjectIndex(modules, undefined, { conditionalCompilation: { projectConstants: {} } });
	const procedures = projectProcedureSignatures(project);
	return modules.flatMap((mod) => analyzeVbaModuleSource({
		source: mod.source, moduleName: mod.moduleName, moduleKind: effectiveModuleKind(mod), host: oracleCase.host,
		...projectAnalysisOptionsForModule(project, mod.moduleName, procedures),
	}).diagnostics.filter((diagnostic) => diagnostic.severity === 'error').map((diagnostic) => `${mod.moduleName}: ${diagnostic.code}: ${diagnostic.message}`));
}

describe('a document module named bare, as a value (issue #1494)', () => {
	it('has every measured case', () => {
		expect(measured).toHaveLength(28);
	});

	it.each(measured.map((oracleCase) => [oracleCase.id, oracleCase] as const))('%s agrees with the VBE', (_id, oracleCase) => {
		const found = errors(oracleCase);
		if (oracleCase.expected === 'accepted') {
			expect(found).toEqual([]);
		} else {
			expect(found).toHaveLength(1);
			expect(found[0]).toMatch(/: (assignment-object-type-mismatch|argument-type-mismatch): .*Run-time error '13'/);
		}
	});

	const project = (host: string, modules: OracleFixtureModule[], body: string[]): OracleCase => ({
		id: 'inline', expected: 'accepted', host,
		modules: [...modules, { name: 'Module1', type: 'standard', source: ['Option Explicit', 'Public Sub Main()', ...body.map((line) => `    ${line}`), 'End Sub', ''].join('\r\n') }],
	});

	it('judges a variable of the document type by the host class too', () => {
		expect(errors(project('excel', [{ name: 'Sheet1', type: 'document', source: '' }], ['Dim s As Sheet1', 'Dim w As Worksheet', 'Set w = s']))).toEqual([]);
		expect(errors(project('excel', [{ name: 'Sheet1', type: 'document', source: '' }], ['Dim s As Sheet1', 'Dim b As Workbook', 'Set b = s']))).toHaveLength(1);
		expect(errors(project('access', [{ name: 'Form_Orders', type: 'document', source: '' }], ['Dim f As Form_Orders', 'Dim g As Form', 'Set g = f']))).toEqual([]);
	});

	it('leaves a module whose class is another host\'s unjudged', () => {
		// A worksheet module in a Word project is no Word class.
		expect(errors(project('word', [{ name: 'Sheet1', type: 'document', source: '' }], ['Dim r As Range', 'Set r = Sheet1']))).toEqual([]);
	});

	it('lets a module-level variable of the same name take precedence', () => {
		const modules: OracleFixtureModule[] = [
			{ name: 'Sheet1', type: 'document', source: '' },
			{ name: 'Module2', type: 'standard', source: 'Option Explicit\r\nPublic Sheet1 As Range\r\n' },
		];
		expect(errors(project('excel', modules, ['Dim r As Range', 'Set r = Sheet1']))).toEqual([]);
	});
});

describe('Extract Variable on a document module (issue #1494)', () => {
	function extracted(host: 'excel' | 'access', documentName: string): string {
		const source = ['Option Explicit', 'Public Sub Main()', `    Debug.Print ${documentName}.Name`, 'End Sub', ''].join('\r\n');
		const modules = [{ moduleName: documentName, type: 'document', source: '' }, { moduleName: 'Module1', type: 'standard', source }];
		const index = buildVbaProjectIndex(modules, undefined, { conditionalCompilation: { projectConstants: {} } });
		const start = source.indexOf(documentName);
		const result = extractVariable({
			source,
			span: { start, end: start + documentName.length },
			typeContext: {
				moduleName: 'Module1',
				moduleKind: 'standard',
				model: host === 'excel' ? getExcelObjectModel() : getAccessObjectModel(),
				projectClassMembers: index.projectMemberSurfaces('Module1', { includeClassValueFacts: false }),
				projectVisibleSymbols: index.visibleIdentifierSymbols('Module1'),
			},
		});
		if (!result.ok) {
			throw new Error(`refused: ${result.reason}`);
		}
		return applyVbaTextEdits(source, result.edits);
	}

	it('declares the module\'s own type and assigns it with Set', () => {
		const text = extracted('excel', 'Sheet1');
		expect(text).toMatch(/Dim (\w+) As Sheet1\r\n\s+Set \1 = Sheet1\r\n/);
	});
});
