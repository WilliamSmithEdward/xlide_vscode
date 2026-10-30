// Issue #1530, measured through the VBE oracle in Excel, Word, PowerPoint and
// Access 16.0 (build 20527). A UserForm or document module sources only the
// Events its code declares in an active arm, so `WithEvents ... As` one with
// none is "Object does not source automation events", bracketed or not. A
// standard module's name is "A module is not a valid type" wherever a type
// goes, and wins over another module's Public Type and VBA's Collection. The
// cases are the issue1530_* fixtures in syntax_corpus/oracle/vbe_oracle_cases.json.

import * as fs from 'node:fs';
import * as path from 'node:path';
import { describe, expect, it } from 'vitest';
import { analyzeVbaModuleSource } from '../../../src/vbaModuleAnalysis';
import { buildVbaProjectIndex, effectiveModuleKind, projectAnalysisOptionsForModule, projectProcedureSignatures } from '../../../src/vbaProjectAnalysis';
import { oracleModuleInput, type OracleFixtureModule } from '../../helpers/oracleModuleInput';

interface OracleCase { id: string; expected: 'accepted' | 'rejected'; host?: string; modules: OracleFixtureModule[] }

const corpus: { cases: OracleCase[] } = JSON.parse(
	fs.readFileSync(path.join(process.cwd(), 'syntax_corpus', 'oracle', 'vbe_oracle_cases.json'), 'utf8'),
);
const measured = corpus.cases.filter((oracleCase) => oracleCase.id.startsWith('issue1530_'));

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

describe('event sources and modules as types (issue #1530)', () => {
	it('has every measured case', () => {
		expect(measured).toHaveLength(108);
	});

	it.each(measured.map((oracleCase) => [oracleCase.id, oracleCase] as const))('%s agrees with the VBE', (_id, oracleCase) => {
		const found = errors(oracleCase);
		if (oracleCase.expected === 'accepted') {
			expect(found).toEqual([]);
		} else {
			// One report, with the VBE's own reason.
			expect(found).toHaveLength(1);
			expect(found[0]).toMatch(oracleCase.id.includes('_module_type_') ? /A module is not a valid type/ : /does not source automation events/);
		}
	});

	it('names the module a WithEvents variable or a Dim is declared As', () => {
		const modules: OracleCase['modules'] = [
			{ name: 'Helpers', type: 'standard', source: 'Option Explicit\r\nSub P()\r\nEnd Sub\r\n' },
			{ name: 'Holder', type: 'class', source: 'Option Explicit\r\nPrivate WithEvents h As Helpers\r\nPrivate m As Helpers\r\n' },
		];
		expect(errors({ modules })).toEqual([
			expect.stringContaining("WithEvents variable 'h' is declared As Helpers, and Helpers is a standard module of this project."),
			expect.stringContaining("'Helpers' is a standard module of this project, and a module is not a type."),
		]);
	});

	it('keeps a UserForm whose code declares an Event a valid source', () => {
		const modules: OracleCase['modules'] = [
			{ name: 'Picker', type: 'userform', source: 'Option Explicit\r\nPublic Event Picked(ByVal index As Long)\r\n' },
			{ name: 'Holder', type: 'class', source: 'Option Explicit\r\nPrivate WithEvents p As Picker\r\nPrivate Sub p_Picked(ByVal index As Long)\r\n    Debug.Print index\r\nEnd Sub\r\n' },
		];
		expect(errors({ modules })).toEqual([]);
	});
});
