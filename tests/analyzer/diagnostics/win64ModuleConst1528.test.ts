// Issue #1528, measured through the VBE oracle in 64-bit Excel, Word,
// PowerPoint and Access 16.0 (build 20527). A module `#Const Win64 = 0`
// turns only `#If Win64` off: 64-bit Office still refuses a Declare without
// PtrSafe, a LongPtr or VarPtr into a Long, and AddressOf into a ByVal Long.
// The rules read the host's compiler constants, so they still report there.
// The cases are the issue1528_* fixtures in syntax_corpus/oracle/vbe_oracle_cases.json.

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
const measured = corpus.cases.filter((oracleCase) => oracleCase.id.startsWith('issue1528_'));

function errorCodes(oracleCase: OracleCase): string[] {
	const modules = oracleCase.modules.map(oracleModuleInput);
	const project = buildVbaProjectIndex(modules, undefined, { conditionalCompilation: { projectConstants: {} } });
	const procedures = projectProcedureSignatures(project);
	return modules.flatMap((mod) => analyzeVbaModuleSource({
		source: mod.source, moduleName: mod.moduleName, moduleKind: effectiveModuleKind(mod), host: oracleCase.host,
		...projectAnalysisOptionsForModule(project, mod.moduleName, procedures),
	}).diagnostics.filter((diagnostic) => diagnostic.severity === 'error').map((diagnostic) => diagnostic.code));
}

const ruleFor = (id: string): string => /addressof/.test(id) ? 'addressof-misuse'
	: /varptr|longptr_into/.test(id) ? 'longlong-narrowing'
		: 'declare-missing-ptrsafe';

describe('64-bit rules under a module #Const Win64 = 0 (issue #1528)', () => {
	it('has every measured case', () => {
		expect(measured).toHaveLength(48);
	});

	it.each(measured.map((oracleCase) => [oracleCase.id, oracleCase] as const))('%s agrees with the VBE', (_id, oracleCase) => {
		const codes = errorCodes(oracleCase);
		if (oracleCase.expected === 'accepted') {
			expect(codes).toEqual([]);
		} else {
			expect(codes).toEqual([ruleFor(oracleCase.id)]);
		}
	});
});
