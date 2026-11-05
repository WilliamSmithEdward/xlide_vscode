// Issue #1621. A referenced VBA project XLIDE did not read (a network path, a
// missing, locked or protected file) may supply any bare name, and the VBE
// accepts its names in every host. The oracle's referenced-project cases,
// measured in Excel, Word, PowerPoint and Access 16.0 (build 20527), compile
// with ProbeLibrary; read without it, XLIDE gave 48 false "Variable not
// defined" errors across 20 of them. Behind an unread project a name XLIDE
// cannot place is now left unjudged.

import * as fs from 'node:fs';
import * as path from 'node:path';
import { describe, expect, it } from 'vitest';
import { analyzeVbaModuleSource } from '../../../src/vbaModuleAnalysis';
import { buildVbaProjectIndex, effectiveModuleKind, projectAnalysisOptionsForModule, projectProcedureSignatures } from '../../../src/vbaProjectAnalysis';
import { oracleModuleInput, oracleProjectReferences, type OracleFixtureModule } from '../../helpers/oracleModuleInput';

interface OracleCase {
	id: string;
	expected: 'accepted' | 'rejected';
	host?: string;
	projectReference?: { name: string; modules: OracleFixtureModule[] };
	modules: OracleFixtureModule[];
}

const corpus: { cases: OracleCase[] } = JSON.parse(
	fs.readFileSync(path.join(process.cwd(), 'syntax_corpus', 'oracle', 'vbe_oracle_cases.json'), 'utf8'),
);
const accepted = corpus.cases.filter((oracleCase) => oracleCase.projectReference && oracleCase.expected === 'accepted');

function errors(oracleCase: OracleCase, read: boolean): string[] {
	const modules = oracleCase.modules.map(oracleModuleInput);
	const project = buildVbaProjectIndex(modules, undefined, { conditionalCompilation: { projectConstants: {} } });
	const procedures = projectProcedureSignatures(project);
	const references = oracleProjectReferences(oracleCase as Parameters<typeof oracleProjectReferences>[0]);
	const options = read ? references : { ...references, referencedProjectSurfaces: [] };
	return modules.flatMap((mod) => analyzeVbaModuleSource({
		source: mod.source, moduleName: mod.moduleName, moduleKind: effectiveModuleKind(mod), host: oracleCase.host,
		...options, ...projectAnalysisOptionsForModule(project, mod.moduleName, procedures),
	}).diagnostics.filter((diagnostic) => diagnostic.severity === 'error').map((diagnostic) => diagnostic.message));
}

function run(code: string, references: Record<string, unknown>): string[] {
	const source = `Option Explicit\r\nPublic Sub Main()\r\n    Dim x As Variant\r\n    ${code}\r\nEnd Sub\r\n`;
	const project = buildVbaProjectIndex([{ moduleName: 'Module1', source }], undefined, { conditionalCompilation: { projectConstants: {} } });
	return analyzeVbaModuleSource({
		source, moduleName: 'Module1', moduleKind: 'standard', host: 'excel', projectName: 'VBAProject', ...references,
		...projectAnalysisOptionsForModule(project, 'Module1', projectProcedureSignatures(project)),
	}).diagnostics.filter((diagnostic) => diagnostic.severity === 'error').map((diagnostic) => diagnostic.message);
}

describe('a bare name behind a referenced project XLIDE did not read (issue #1621)', () => {
	it('is left unjudged in every accepted referenced-project case, in all four hosts', () => {
		const hosts = new Set(accepted.map((oracleCase) => oracleCase.host ?? 'excel'));
		expect([...hosts].sort()).toEqual(['access', 'excel', 'powerpoint', 'word']);
		for (const oracleCase of accepted) {
			expect(errors(oracleCase, true), `${oracleCase.id}, read`).toEqual([]);
			expect(errors(oracleCase, false), `${oracleCase.id}, unread`).toEqual([]);
		}
	});

	it('leaves a bare name, a call and the project\'s own name unjudged behind an unread project', () => {
		const unread = { referencedLibraries: ['stdole', 'Office', 'Unread'], referencedProjects: ['Unread'], referencedProjectSurfaces: [] };
		expect(run('x = Widget', unread)).toEqual([]);
		expect(run('Call Helper(1)', unread)).toEqual([]);
		expect(run('Helper 1', unread)).toEqual([]);
		expect(run('x = Unread.Service.Value()', unread)).toEqual([]);
	});

	it('still reports a name nothing could supply where every reference is known', () => {
		const known = { referencedLibraries: ['stdole', 'Office'] };
		expect(run('x = Nonexistent', known)).toEqual([expect.stringMatching(/^Variable not defined: 'Nonexistent'/)]);
		expect(run('Helper 1', known)).toEqual([expect.stringMatching(/^Sub or Function not defined/)]);
	});

	it('still reports `v = Widget` when only a read ProbeLibrary has the name', () => {
		const oracleCase = accepted.find((candidate) => candidate.projectReference?.name === 'ProbeLibrary')!;
		const references = oracleProjectReferences(oracleCase as Parameters<typeof oracleProjectReferences>[0]);
		const messages = run('x = Widget', references);
		expect(messages.length).toBeGreaterThan(0);
	});
});
