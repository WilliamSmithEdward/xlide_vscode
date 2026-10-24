// A document module reached from another module, through its code name or a
// variable of its type, is the same closed class as its own `Me`: the host
// object's members, the module's own and the controls on it, and nothing
// else. Measured in Excel, Word and Access 16.0 (build 20430) through the
// VBE oracle; the document_member_* fixtures in vbe_oracle_cases.json hold
// every verdict, and this test holds the analyzer to each one. PowerPoint has
// no document modules.

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { analyzeVbaModuleSource } from '../../../src/vbaModuleAnalysis';
import {
	buildVbaProjectIndex,
	effectiveModuleKind,
	projectAnalysisOptionsForModule,
	projectProcedureSignatures,
	type VbaProjectModuleInput,
} from '../../../src/vbaProjectAnalysis';
import { byCode } from '../../helpers/diagnostics';
import { oracleModuleInput, type OracleFixtureModule } from '../../helpers/oracleModuleInput';
import { analyzeProjectModule } from './helpers';

interface Fixture {
	id: string;
	host?: string;
	expected: 'accepted' | 'rejected';
	modules: OracleFixtureModule[];
}

const fixtures: Fixture[] = (JSON.parse(readFileSync(join(process.cwd(), 'syntax_corpus', 'oracle', 'vbe_oracle_cases.json'), 'utf8')) as { cases: Fixture[] })
	.cases.filter((kase) => kase.id.startsWith('document_member_'));

/** An Access form's module reaches the analyzer as the extension reads it from the database. */
function moduleInput(fixture: Fixture, mod: OracleFixtureModule): VbaProjectModuleInput {
	const input = oracleModuleInput(mod);
	return fixture.host === 'access' && mod.type === 'document' ? { ...input, type: 'accessform' } : input;
}

function memberNotFound(fixture: Fixture): string[] {
	const modules = fixture.modules.map((mod) => moduleInput(fixture, mod));
	const project = buildVbaProjectIndex(modules);
	const procedures = projectProcedureSignatures(project);
	return modules.flatMap((mod) => byCode(analyzeVbaModuleSource({
		source: mod.source,
		moduleName: mod.moduleName,
		moduleKind: effectiveModuleKind(mod),
		...projectAnalysisOptionsForModule(project, mod.moduleName, procedures),
		...(fixture.host ? { host: fixture.host } : {}),
	}).diagnostics, 'member-not-found').map((diagnostic) => diagnostic.message));
}

describe('document module members against the measured fixtures', () => {
	it('has fixtures for every host with document modules', () => {
		expect(new Set(fixtures.map((fixture) => fixture.host ?? 'excel'))).toEqual(new Set(['excel', 'word', 'access']));
	});

	// A renamed workbook or document module (BookCode, DocCode) is known by
	// the document type the host reads from its VB_Base, so it is checked
	// like ThisWorkbook and ThisDocument (#1457).
	it.each(fixtures.map((fixture) => [fixture.id, fixture] as const))('%s', (_id, fixture) => {
		expect(memberNotFound(fixture)).toHaveLength(fixture.expected === 'rejected' ? 1 : 0);
	});
});

describe('document module members elsewhere', () => {
	const HELLO = 'Option Explicit\nPublic Sub Hello()\nEnd Sub\n';
	const caller = (body: string) => `Option Explicit\nPublic Sub T()\n${body}\nEnd Sub\n`;

	function findings(host: string, documentModule: Record<string, unknown>, body: string): string[] {
		return byCode(analyzeProjectModule(caller(body), [
			{ moduleName: 'Module1', source: '' },
			{ source: HELLO, moduleKind: 'document', ...documentModule } as never,
		], 'Module1', { host } as never), 'member-not-found').map((diagnostic) => diagnostic.message);
	}

	it('names the module and the member', () => {
		expect(findings('excel', { moduleName: 'ThisWorkbook' }, '    Dim w As ThisWorkbook\n    Debug.Print w.UsedRange'))
			.toEqual([expect.stringContaining("'ThisWorkbook.UsedRange'")]);
	});

	it('leaves a sheet unchecked while its controls are unknown', () => {
		expect(findings('excel', { moduleName: 'Sheet1' }, '    Dim s As Sheet1\n    Debug.Print s.FullName\n    Debug.Print Sheet1.FullName')).toEqual([]);
	});

	it('leaves a plain Workbook or Document variable unchecked', () => {
		expect(findings('excel', { moduleName: 'ThisWorkbook' }, '    Dim w As Workbook\n    Debug.Print w.UsedRange')).toEqual([]);
		expect(findings('word', { moduleName: 'ThisDocument' }, '    Dim d As Document\n    Debug.Print d.UsedRange')).toEqual([]);
	});

	it("checks a module only against its own host's class", () => {
		// A Word document module named ThisWorkbook is not a workbook, and
		// an Excel one named ThisDocument is not a Word document.
		expect(findings('word', { moduleName: 'ThisWorkbook' }, '    Dim w As ThisWorkbook\n    Debug.Print w.Nope')).toEqual([]);
		expect(findings('excel', { moduleName: 'ThisDocument' }, '    Dim d As ThisDocument\n    Debug.Print d.Nope')).toEqual([]);
	});
});
