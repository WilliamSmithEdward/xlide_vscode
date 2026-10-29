// Measured through the VBE oracle in Excel, Word, PowerPoint and Access 16.0
// (build 20527), every host alike. A UserForm or a document module (a
// workbook, sheet, Word document, slide, Access form or report) sources only
// the Events its code declares: with none, or one only in an inactive #If
// arm, `WithEvents x As` it is "Object does not source automation events",
// a control on it or the host's own events notwithstanding. And a standard
// module's name is no type: in a Dim, a procedure's parameter or return, a
// Type member or a WithEvents it is "A module is not a valid type", even
// where another module's Type of that name, VBA's Collection or a host type
// such as Range would otherwise answer. A Type or Enum of the declaring
// module's own wins over the module, and a Const is "Expected: type name"
// alone. A bracketed name is the same name. The cases are the
// withevents_source_* and module_as_type_* fixtures in
// syntax_corpus/oracle/vbe_oracle_cases.json.

import * as fs from 'node:fs';
import * as path from 'node:path';
import { describe, expect, it } from 'vitest';
import { analyzeVbaModuleSource } from '../../../src/vbaModuleAnalysis';
import { buildVbaProjectIndex, effectiveModuleKind, projectAnalysisOptionsForModule, projectProcedureSignatures } from '../../../src/vbaProjectAnalysis';
import { oracleModuleInput, type OracleFixtureModule } from '../../helpers/oracleModuleInput';

interface OracleCase { id: string; description: string; expected: 'accepted' | 'rejected'; host?: string; modules: OracleFixtureModule[] }

const corpus: { cases: OracleCase[] } = JSON.parse(
	fs.readFileSync(path.join(process.cwd(), 'syntax_corpus', 'oracle', 'vbe_oracle_cases.json'), 'utf8'),
);
const sources = corpus.cases.filter((oracleCase) => oracleCase.id.startsWith('withevents_source_'));
const moduleTypes = corpus.cases.filter((oracleCase) => oracleCase.id.startsWith('module_as_type_'));

function errors(oracleCase: Pick<OracleCase, 'host' | 'modules'>): string[] {
	const modules = oracleCase.modules.map(oracleModuleInput);
	const project = buildVbaProjectIndex(modules, undefined, { conditionalCompilation: { projectConstants: {} } });
	const procedures = projectProcedureSignatures(project);
	return modules.flatMap((mod) => analyzeVbaModuleSource({
		source: mod.source, moduleName: mod.moduleName, moduleKind: effectiveModuleKind(mod), host: oracleCase.host,
		designerClass: mod.designerClass, implicitMembers: mod.implicitMembers,
		...projectAnalysisOptionsForModule(project, mod.moduleName, procedures),
	}).diagnostics.filter((diagnostic) => diagnostic.severity === 'error').map((diagnostic) => `${mod.moduleName}: ${diagnostic.code}: ${diagnostic.message}`));
}

describe('event sources and standard modules as types', () => {
	it('has every measured case', () => {
		expect(sources).toHaveLength(48);
		expect(moduleTypes).toHaveLength(108);
	});

	it.each(sources.map((oracleCase) => [oracleCase.id, oracleCase] as const))('%s agrees with the VBE', (_id, oracleCase) => {
		expect(errors(oracleCase)).toEqual(oracleCase.expected === 'accepted' ? [] : [
			expect.stringMatching(/^Sink: withevents-declaration: WithEvents variable 'source' is declared As \[?\w+\]?, which does not source automation events\.$/),
		]);
	});

	it.each(moduleTypes.map((oracleCase) => [oracleCase.id, oracleCase] as const))('%s agrees with the VBE', (_id, oracleCase) => {
		const found = errors(oracleCase);
		if (oracleCase.expected === 'accepted') {
			expect(found).toEqual([]);
			return;
		}
		expect(found).toHaveLength(1);
		if (oracleCase.description.includes('Expected: type name')) {
			expect(found[0]).toMatch(/^Module3: const-invalid-type: .* This is a VBE compile error: Expected: type name\.$/);
			return;
		}
		const [, reporter, name] = /^(\w+): invalid-as-type-name: '(\w+)' is a standard module of this project, not a type\. This is a VBE compile error: A module is not a valid type\.$/.exec(found[0]) ?? [];
		// The module that declares the type reports it, and the name is a standard module's.
		expect(oracleCase.modules.find((mod) => mod.name === name)?.type).toBe('standard');
		expect(oracleCase.modules.find((mod) => mod.name === reporter)?.source).toMatch(new RegExp(`As (New )?\\[?${name}\\]?\\b`));
	});

	it('judges a UserForm by the Events its code declares, not by its controls', () => {
		const sink = { name: 'Sink', type: 'class', source: 'Option Explicit\r\nPrivate WithEvents f As UserForm1\r\n' };
		const form = (source: string): OracleFixtureModule => ({ name: 'UserForm1', type: 'userform', source, designerControls: [{ name: 'CommandButton1', progId: 'Forms.CommandButton.1' }] });
		expect(errors({ modules: [form('Option Explicit\r\n'), sink] })).toHaveLength(1);
		expect(errors({ modules: [form('Option Explicit\r\nPublic Event Changed(ByVal value As Long)\r\n'), sink] })).toEqual([]);
	});

	it('leaves a qualified name and a project class alone', () => {
		const modules = (declaration: string): OracleFixtureModule[] => [
			{ name: 'Range', type: 'standard', source: 'Option Explicit\r\n' },
			{ name: 'Holder', type: 'class', source: 'Option Explicit\r\n' },
			{ name: 'Module3', type: 'standard', source: `Option Explicit\r\nPublic Sub T()\r\n    ${declaration}\r\n    Debug.Print 1\r\nEnd Sub\r\n` },
		];
		expect(errors({ host: 'excel', modules: modules('Dim r As Excel.Range') })).toEqual([]);
		expect(errors({ host: 'excel', modules: modules('Dim h As Holder') })).toEqual([]);
		expect(errors({ host: 'excel', modules: modules('Dim h As New Holder') })).toEqual([]);
		expect(errors({ host: 'excel', modules: modules('Dim r As Range') })).toEqual([
			"Module3: invalid-as-type-name: 'Range' is a standard module of this project, not a type. This is a VBE compile error: A module is not a valid type.",
		]);
	});
});
