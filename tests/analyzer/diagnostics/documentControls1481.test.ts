// Issue #1481, measured through the VBE oracle in Word, Excel and PowerPoint
// 16.0 (build 20527) with ActiveX enabled. An ActiveX control on a Word
// document, an Excel sheet or a PowerPoint slide is a member of the module
// behind it: named bare and through Me there, and through the module's name
// or a variable of its type from another module. Its own members are bound
// when the code runs, so `TextBox1.Nope` compiles. The cases are the
// issue1481_* fixtures in syntax_corpus/oracle/vbe_oracle_cases.json, whose
// document modules carry what the engine reads for them.

import * as fs from 'node:fs';
import * as path from 'node:path';
import { describe, expect, it } from 'vitest';
import { readModules } from '../../../src/vba/projectService';
import { analyzeVbaModuleSource } from '../../../src/vbaModuleAnalysis';
import { buildVbaProjectIndex, effectiveModuleKind, projectAnalysisOptionsForModule, projectProcedureSignatures } from '../../../src/vbaProjectAnalysis';
import { oracleModuleInput, type OracleFixtureModule } from '../../helpers/oracleModuleInput';

interface OracleCase { id: string; expected: 'accepted' | 'rejected'; host?: string; modules: OracleFixtureModule[] }

const corpus: { cases: OracleCase[] } = JSON.parse(
	fs.readFileSync(path.join(process.cwd(), 'syntax_corpus', 'oracle', 'vbe_oracle_cases.json'), 'utf8'),
);
const measured = corpus.cases.filter((oracleCase) => oracleCase.id.startsWith('issue1481_'));
const BINARIES = path.join(__dirname, '..', '..', 'fixtures', 'binaries');

function errors(oracleCase: OracleCase): string[] {
	const modules = oracleCase.modules.map(oracleModuleInput);
	const project = buildVbaProjectIndex(modules, undefined, { conditionalCompilation: { projectConstants: {} } });
	const procedures = projectProcedureSignatures(project);
	return modules.flatMap((mod) => analyzeVbaModuleSource({
		source: mod.source, moduleName: mod.moduleName, moduleKind: effectiveModuleKind(mod), host: oracleCase.host,
		designerClass: mod.designerClass, implicitMembers: mod.implicitMembers,
		...projectAnalysisOptionsForModule(project, mod.moduleName, procedures),
	}).diagnostics.filter((diagnostic) => diagnostic.severity === 'error').map((diagnostic) => `${mod.moduleName}: ${diagnostic.code}`));
}

describe('ActiveX controls on a document, a sheet or a slide (issue #1481)', () => {
	it('has every measured case', () => {
		expect(measured).toHaveLength(34);
	});

	it.each(measured.map((oracleCase) => [oracleCase.id, oracleCase] as const))('%s agrees with the VBE', (_id, oracleCase) => {
		const found = errors(oracleCase);
		if (oracleCase.expected === 'accepted') {
			expect(found).toEqual([]);
		} else {
			expect(found.length).toBeGreaterThan(0);
		}
	});
});

describe('the controls a module lists, as the engine reads them (issue #1481)', () => {
	const document = (file: string, name: string) => readModules(path.join(BINARIES, file)).find((mod) => mod.name === name);

	it('gives a Word document module every control Word wrote: inline, floating and in a header', () => {
		const module = document('WordControlsFixture.docm', 'ThisDocument');
		expect(module?.designerClass).toBe('Word.Document');
		expect(module?.implicitMembers).toEqual([
			{ name: 'CommandButton1', type: 'Object' },
			{ name: 'CommandButton2', type: 'Object' },
			{ name: 'CheckBox1', type: 'Object' },
		]);
	});

	it('gives a PowerPoint slide module its control', () => {
		const module = document('PowerPointControlsFixture.pptm', 'Slide1');
		expect(module?.designerClass).toBe('PowerPoint.Slide');
		expect(module?.implicitMembers).toEqual([{ name: 'CommandButton1', type: 'Object' }]);
	});

	it('leaves a document with no control as it was', () => {
		const module = document('WordFixture.docm', 'ThisDocument');
		expect(module?.designerClass).toBeUndefined();
		expect(module?.implicitMembers).toBeUndefined();
	});
});
