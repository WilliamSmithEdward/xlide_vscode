// A Collection assigned whole to a member or an array element: `k.n = c`,
// `a(0) = c`, `Range("A1").Value = c`. Measured in Excel, Word, PowerPoint
// and Access 16.0 (build 20527) through the VBE oracle: into a target the VBE
// knows as a scalar or an enum - a class's Long field or Property Let, a UDT
// field, a typed array's element, a String, enum, Single or Boolean host
// property - it is "Argument not optional"; into a Variant, a Variant
// property or a late-bound receiver it compiles, and fails as it runs. The
// cases are the collection_into_* fixtures in syntax_corpus/oracle/vbe_oracle_cases.json.

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
import { oracleDefaultDocumentModules, oracleModuleInput, oracleProjectReferences, type OracleFixtureModule } from '../../helpers/oracleModuleInput';

interface OracleCase {
	id: string;
	expected: 'accepted' | 'rejected';
	host: string;
	modules: OracleFixtureModule[];
}

const fixtures: OracleCase[] = (JSON.parse(
	fs.readFileSync(path.join(process.cwd(), 'syntax_corpus', 'oracle', 'vbe_oracle_cases.json'), 'utf8'),
) as { cases: OracleCase[] }).cases.filter((kase) => kase.id.startsWith('collection_into_'));

/** Each compile-error code, for the project as the extension reads it. */
function errorCodes(kase: { host: string; modules: OracleFixtureModule[] }): string[] {
	const modules = kase.modules.map(oracleModuleInput);
	const all = [...modules, ...oracleDefaultDocumentModules(kase, kase.modules).map(oracleModuleInput)];
	const project = buildVbaProjectIndex(all, undefined, { conditionalCompilation: { projectConstants: {} } });
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

const procedure = (host: string, ...lines: string[]) => ({
	host,
	modules: [{ name: 'Module1', type: 'standard', source: ['Option Explicit', 'Sub Main()', '    Dim c As New Collection, v As Variant', ...lines.map((line) => '    ' + line), 'End Sub', ''].join('\r\n') }],
});

describe('a Collection into a member or element, against the measured fixtures', () => {
	it('has fixtures for every Office host', () => {
		expect(new Set(fixtures.map((fixture) => fixture.host))).toEqual(new Set(['excel', 'word', 'powerpoint', 'access']));
	});

	it.each(fixtures.map((fixture) => [fixture.id, fixture] as const))('%s', (_id, fixture) => {
		expect(errorCodes(fixture)).toEqual(fixture.expected === 'rejected' ? ['collection-operand'] : []);
	});
});

describe('a Collection into a member or element elsewhere', () => {
	it('still reports a scalar target, a project enum, an indexed setter and an object-holding default, as main did', () => {
		expect(errorCodes(procedure('excel', 'Dim ws As Worksheet', 'If False Then ws.Name = c'))).toEqual(['collection-operand']);
		const kase = (body: string, cls: string) => ({ host: 'excel', modules: [
			{ name: 'Module1', type: 'standard', source: ['Option Explicit', 'Sub Main()', '    Dim c As New Collection, k As New Class1', '    ' + body, 'End Sub', ''].join('\r\n') },
			{ name: 'Class1', type: 'class', source: ['Option Explicit', ...cls.split('\n'), ''].join('\r\n') },
		] });
		expect(errorCodes(kase('k.n = c', 'Public Enum E\n    A\nEnd Enum\nPublic n As E'))).toEqual(['collection-operand']);
		expect(errorCodes(kase('k.P(1) = c', 'Public Property Let P(ByVal i As Long, ByVal x As Long)\nEnd Property'))).toEqual(['collection-operand']);
		// A Public array is refused in a class for its own reason too.
		expect(errorCodes(kase('k.a(0) = c', 'Public a(1) As Long'))).toContain('collection-operand');
		expect(errorCodes(procedure('word', 'Dim p As Paragraph, d As Document', 'If False Then d.Content.Text = p'))).toEqual(['collection-operand']);
	});

	it('takes a two-dimensional Variant array, an array with no As clause, and a With block', () => {
		expect(errorCodes(procedure('excel', 'Dim a(1, 1) As Variant', 'a(0, 1) = c'))).toEqual([]);
		expect(errorCodes(procedure('excel', 'Dim a(1)', 'a(0) = c'))).toEqual([]);
		expect(errorCodes(procedure('excel', 'With Range("A1")', '    .Value = c', 'End With'))).toEqual([]);
	});

	it("reads a local array with no As clause as Variants, though a module variable of its name is typed", () => {
		const kase = { host: 'excel', modules: [{ name: 'Module1', type: 'standard',
			source: ['Option Explicit', 'Dim a As Long', 'Sub Main()', '    Dim c As New Collection', '    Dim a(1)', '    a(0) = c', 'End Sub', ''].join('\r\n') }] };
		expect(errorCodes(kase)).toEqual([]);
	});

	it('reads an array with no As clause under DefInt as Integers', () => {
		const kase = { host: 'excel', modules: [{ name: 'Module1', type: 'standard',
			source: ['DefInt A-Z', 'Option Explicit', 'Sub Main()', '    Dim c As New Collection', '    Dim a(1)', '    a(0) = c', 'End Sub', ''].join('\r\n') }] };
		expect(errorCodes(kase)).toEqual(['collection-operand']);
	});

	it("still reports an assignment through a host function's default member, which compiles", () => {
		// Range("A1") = c compiles (measured in Excel); the default member's
		// type is not read here.
		expect(errorCodes(procedure('excel', 'Range("A1") = c'))).toEqual(['collection-operand']);
	});

	it('still reads a comparison, or an operator, as before', () => {
		expect(errorCodes(procedure('excel', 'Dim r As Range', 'If False Then v = (r.Value = c)'))).toEqual(['collection-operand']);
		expect(errorCodes(procedure('excel', 'Range("A1").Value = c + 1'))).toEqual(['collection-operand']);
	});

	it('takes an element of the Collection, which has a value', () => {
		expect(errorCodes(procedure('excel', 'c.Add 1', 'Range("A1").Value = c(1)'))).toEqual([]);
		expect(errorCodes(procedure('excel', 'Dim a(1) As Long', 'c.Add 1', 'a(0) = c(1)'))).toEqual([]);
	});

	it('reads a Let and a one-line If the same way', () => {
		expect(errorCodes(procedure('excel', 'Let Range("A1").Value = c'))).toEqual([]);
		expect(errorCodes(procedure('excel', 'If v Then Range("A1").Value = c'))).toEqual([]);
		expect(errorCodes(procedure('excel', 'Dim a(1) As Long', 'If v Then a(0) = c'))).toEqual(['collection-operand']);
	});
});
