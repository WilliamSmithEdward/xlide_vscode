// Beep and Reset take no argument, Randomize one optional, Load and Unload
// exactly one, as statements and through Call. Measured in Excel, Word,
// PowerPoint and Access 16.0 (build 20527) through the VBE oracle: one too
// many is "Wrong number of arguments or invalid property assignment", and
// Load or Unload with none "Argument not optional". The cases are the
// statement_arguments_* fixtures in syntax_corpus/oracle/vbe_oracle_cases.json.

import * as fs from 'node:fs';
import * as path from 'node:path';
import { describe, expect, it } from 'vitest';
import { analyzeVbaModuleSource } from '../../../src/vbaModuleAnalysis';
import {
	buildVbaProjectIndex,
	effectiveModuleKind,
	projectAnalysisOptionsForModule,
	projectProcedureSignatures,
} from '../../../src/vbaProjectAnalysis';
import { oracleModuleInput, oracleProjectReferences, type OracleFixtureModule } from '../../helpers/oracleModuleInput';

interface OracleCase {
	id: string;
	expected: 'accepted' | 'rejected';
	host: string;
	modules: OracleFixtureModule[];
}

const fixtures: OracleCase[] = (JSON.parse(
	fs.readFileSync(path.join(process.cwd(), 'syntax_corpus', 'oracle', 'vbe_oracle_cases.json'), 'utf8'),
) as { cases: OracleCase[] }).cases.filter((kase) => kase.id.startsWith('statement_arguments_'));

/** Each compile-error code, for the project as the extension reads it. */
function errorCodes(kase: { host: string; modules: OracleFixtureModule[] }): string[] {
	const modules = kase.modules.map(oracleModuleInput);
	const project = buildVbaProjectIndex(modules, undefined, { conditionalCompilation: { projectConstants: {} } });
	const procedures = projectProcedureSignatures(project);
	return modules.flatMap((mod) => analyzeVbaModuleSource({
		source: mod.source,
		moduleName: mod.moduleName,
		moduleKind: effectiveModuleKind(mod),
		host: kase.host,
		...oracleProjectReferences({ host: kase.host }),
		...projectAnalysisOptionsForModule(project, mod.moduleName, procedures),
	}).diagnostics.filter((diagnostic) => diagnostic.severity === 'error').map((diagnostic) => diagnostic.code ?? ''));
}

const procedure = (...lines: string[]) => ({
	host: 'excel',
	modules: [{ name: 'Module1', type: 'standard', source: ['Option Explicit', 'Sub Main()', '    Dim c As New Collection', ...lines.map((line) => '    ' + line), 'End Sub', ''].join('\r\n') }],
});

describe('statement argument counts against the measured fixtures', () => {
	it('has fixtures for every Office host', () => {
		expect(new Set(fixtures.map((fixture) => fixture.host))).toEqual(new Set(['excel', 'word', 'powerpoint', 'access']));
	});

	it.each(fixtures.map((fixture) => [fixture.id, fixture] as const))('%s', (_id, fixture) => {
		expect(errorCodes(fixture)).toEqual(fixture.expected === 'accepted' ? [] : ['argument-count']);
	});
});

describe('statement argument counts elsewhere', () => {
	it('reads a one-line If branch and a line after a colon', () => {
		expect(errorCodes(procedure('If True Then Beep 1'))).toEqual(['argument-count']);
		expect(errorCodes(procedure('Beep: Reset 1'))).toEqual(['argument-count']);
	});

	it('takes Load and Unload of a UserForm with one argument', () => {
		const kase = { host: 'excel', modules: [
			{ name: 'Module1', type: 'standard', source: ['Option Explicit', 'Sub Main()', '    Load Form1', '    Unload Form1', 'End Sub', ''].join('\r\n') },
			{ name: 'Form1', type: 'userform', source: ['Option Explicit', 'Private Sub UserForm_Click()', '    Unload Me', 'End Sub', ''].join('\r\n') },
		] };
		expect(errorCodes(kase)).toEqual([]);
	});

	it('leaves a label named Reset alone', () => {
		expect(errorCodes(procedure('GoTo Reset', 'Reset:'))).toEqual([]);
	});

	it("ends a one-line If's branch at its Else and a statement at its colon", () => {
		expect(errorCodes(procedure('If True Then Beep Else MsgBox "a"'))).toEqual([]);
		expect(errorCodes(procedure('If True Then Beep: Beep Else Reset'))).toEqual([]);
		expect(errorCodes(procedure('If True Then Beep: Beep 1 Else Beep'))).toEqual(['argument-count']);
		expect(errorCodes(procedure('If True Then Beep Else Beep 1'))).toEqual(['argument-count']);
		// A Single's type suffix and then Else: the Else isn't a member name.
		expect(errorCodes(procedure('Dim x As Single', 'If True Then Beep: Randomize x! Else MsgBox "a", 1'))).toEqual([]);
		expect(errorCodes(procedure('Dim x As Single', 'If True Then Randomize x! Else Beep 1'))).toEqual(['argument-count']);
	});

	it("reads a Then branch without its Else for every rule, as the overflow one", () => {
		expect(errorCodes(procedure('Dim x As Integer', 'If True Then x = 40000 Else x = 1'))).toContain('arithmetic-overflow');
	});

	it('reads a label at the start of a line after a lone CR or a line number', () => {
		const lines = (...body: string[]) => ({ host: 'excel', modules: [{ name: 'Module1', type: 'standard', source: ['Option Explicit', 'Sub Main()', ...body, 'End Sub', ''].join('\r') }] });
		expect(errorCodes(lines('    GoTo Kill', 'Kill:'))).toEqual([]);
		expect(errorCodes(procedure('GoTo Load', '10 Load: Beep'))).toEqual([]);
	});

	it('reads Call with no argument list, a parenthesized first argument and a statement after a colon', () => {
		expect(errorCodes(procedure('Dim o As Object', 'Call Unload'))).toEqual(['argument-count']);
		expect(errorCodes(procedure('Dim o As Object', 'Load (o), o'))).toEqual(['argument-count']);
		expect(errorCodes(procedure('Dim o As Object', 'Load (o)'))).toEqual([]);
		expect(errorCodes(procedure('Beep: Unload: Beep'))).toEqual(['argument-count']);
		expect(errorCodes(procedure('Kill ("a")'))).toEqual([]);
	});
});
