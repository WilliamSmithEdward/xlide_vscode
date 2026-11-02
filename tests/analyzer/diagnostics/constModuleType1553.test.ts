// Issue #1553, measured through the VBE oracle in Excel, Word, PowerPoint and
// Access 16.0 (build 20527). A Const declared As a type VBA does not own is one
// compile error: "Expected: type name" at module level, "Syntax error" inside
// a procedure, and "Invalid data type for constant" for Object at either
// level. A standard module's name is no exception: the Const reports it, and
// "A module is not a valid type" does not. The cases are the issue1553_*
// fixtures in syntax_corpus/oracle/vbe_oracle_cases.json.

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
const measured = corpus.cases.filter((oracleCase) => oracleCase.id.startsWith('issue1553_'));

function errors(oracleCase: Pick<OracleCase, 'host' | 'modules'>): string[] {
	const modules = oracleCase.modules.map(oracleModuleInput);
	const project = buildVbaProjectIndex(modules, undefined, { conditionalCompilation: { projectConstants: {} } });
	const procedures = projectProcedureSignatures(project);
	return modules.flatMap((mod) => analyzeVbaModuleSource({
		source: mod.source, moduleName: mod.moduleName, moduleKind: effectiveModuleKind(mod), host: oracleCase.host,
		...projectAnalysisOptionsForModule(project, mod.moduleName, procedures),
	}).diagnostics.filter((diagnostic) => diagnostic.severity === 'error').map((diagnostic) => `${diagnostic.code}: ${diagnostic.message}`));
}

/** The VBE's message for a refused case, from its shape. */
function vbeMessage(id: string): RegExp {
	if (id.includes('object')) {
		return /Invalid data type for constant/;
	}
	return id.includes('procedure_') ? /Syntax error/ : /Expected: type name/;
}

describe('a Const As a type VBA does not own (issue #1553)', () => {
	it('has every measured case', () => {
		// Ten shapes refused and two accepted, in each of the four hosts.
		expect(measured).toHaveLength(48);
	});

	it.each(measured.map((oracleCase) => [oracleCase.id, oracleCase] as const))('%s agrees with the VBE', (_id, oracleCase) => {
		const found = errors(oracleCase);
		if (oracleCase.expected === 'accepted') {
			expect(found).toEqual([]);
		} else {
			// One report, with the VBE's own reason.
			expect(found).toHaveLength(1);
			expect(found[0]).toMatch(/^const-invalid-type: /);
			expect(found[0]).toMatch(vbeMessage(oracleCase.id));
		}
	});
});
