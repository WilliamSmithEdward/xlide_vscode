// `Date = #1/1/2020#` is VBA's Date statement, which sets the system date,
// as `Time = t` sets the time. Measured in Excel, Word, PowerPoint and Access
// 16.0 (build 20527) through the VBE oracle: it compiles with Option Explicit,
// and `Let Date = d` is a Syntax error. The cases are the date_statement_*
// fixtures in syntax_corpus/oracle/vbe_oracle_cases.json.

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
) as { cases: OracleCase[] }).cases.filter((kase) => kase.id.startsWith('date_statement_'));

/** Each error code, with the project known as the extension knows it, which turns on the undeclared-name check. */
function errorCodes(source: string, host = 'excel'): string[] {
	const modules = [oracleModuleInput({ name: 'Module1', type: 'standard', source })];
	const project = buildVbaProjectIndex(modules, undefined, { conditionalCompilation: { projectConstants: {} } });
	const procedures = projectProcedureSignatures(project);
	return modules.flatMap((mod) => analyzeVbaModuleSource({
		source: mod.source,
		moduleName: mod.moduleName,
		moduleKind: effectiveModuleKind(mod),
		host,
		...oracleProjectReferences({ host }),
		...projectAnalysisOptionsForModule(project, mod.moduleName, procedures),
	}).diagnostics.filter((diagnostic) => diagnostic.severity === 'error').map((diagnostic) => diagnostic.code ?? ''));
}

const procedure = (...lines: string[]) => ['Option Explicit', 'Sub Main()', '    Dim d As Date', ...lines.map((line) => '    ' + line), 'End Sub', ''].join('\r\n');

describe('the Date statement against the measured fixtures', () => {
	it('has fixtures for every Office host', () => {
		expect(new Set(fixtures.map((fixture) => fixture.host))).toEqual(new Set(['excel', 'word', 'powerpoint', 'access']));
	});

	it.each(fixtures.map((fixture) => [fixture.id, fixture] as const))('%s', (_id, fixture) => {
		const codes = errorCodes(fixture.modules[0].source, fixture.host);
		expect(codes).toEqual(fixture.expected === 'accepted' ? [] : ['undeclared-variable']);
	});
});

describe('the Date statement elsewhere', () => {
	it('is a statement in any case, after a colon, a label or a line number, and continued', () => {
		expect(errorCodes(procedure('date = d'))).toEqual([]);
		expect(errorCodes(procedure('d = Now: Date = d'))).toEqual([]);
		expect(errorCodes(procedure('Here: Date = d'))).toEqual([]);
		expect(errorCodes(procedure('10 Date = d'))).toEqual([]);
		expect(errorCodes(procedure('Date _', '    = d'))).toEqual([]);
	});

	it('takes any value the VBE compiles', () => {
		expect(errorCodes(procedure('Date = #1/1/2020#'))).toEqual([]);
		expect(errorCodes(procedure('Date = Date'))).toEqual([]);
		expect(errorCodes(procedure('Date() = d'))).toEqual([]);
		expect(errorCodes(procedure('Date = "let me"'))).toEqual([]);
	});

	it('still reads its value', () => {
		expect(errorCodes(procedure('Date = nope'))).toEqual(['undeclared-variable']);
	});

	it('leaves Let Date after a label or line number, and Set Date, to the report, as the VBE refuses them', () => {
		expect(errorCodes(procedure('Here: Let Date = d'))).toEqual(['undeclared-variable']);
		expect(errorCodes(procedure('10 Let Date = d'))).toEqual(['undeclared-variable']);
		expect(errorCodes(procedure('Dim o As Object', 'Set Date = o'))).toEqual(['undeclared-variable']);
	});
});
