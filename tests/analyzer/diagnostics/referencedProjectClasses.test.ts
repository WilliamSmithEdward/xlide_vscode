// The classes of a referenced VBA project, measured through the VBE oracle in
// Excel, Word, PowerPoint and Access 16.0 (build 20527), every host alike:
// ProbeLibrary's class Widget is exposed (VB_Exposed, Instancing
// PublicNotCreatable), Solo is exposed with a default instance and Thing is
// private. The cases are the referenced_class_* fixtures in
// syntax_corpus/oracle/vbe_oracle_cases.json.

import * as fs from 'node:fs';
import * as path from 'node:path';
import { describe, expect, it } from 'vitest';
import { referencedProjectSurface, type ReferencedProjectModuleSource, type ReferencedProjectSurface } from '../../../src/analyzer/symbols/referencedProjects';
import { analyzeVbaModuleSource } from '../../../src/vbaModuleAnalysis';
import { buildVbaProjectIndex, effectiveModuleKind, projectAnalysisOptionsForModule, projectProcedureSignatures } from '../../../src/vbaProjectAnalysis';
import { oracleModuleInput, oracleProjectReferences, type OracleFixtureModule } from '../../helpers/oracleModuleInput';

interface OracleCase {
	id: string;
	expected: 'accepted' | 'rejected';
	description: string;
	host?: string;
	projectReference: { name: string; modules: OracleFixtureModule[] };
	modules: OracleFixtureModule[];
}

const corpus: { cases: OracleCase[] } = JSON.parse(
	fs.readFileSync(path.join(process.cwd(), 'syntax_corpus', 'oracle', 'vbe_oracle_cases.json'), 'utf8'),
);
const measured = corpus.cases.filter((oracleCase) => oracleCase.id.startsWith('referenced_class_'));
const byId = (id: string): OracleCase => measured.find((oracleCase) => oracleCase.id === id)!;

/**
 * Where XLIDE cannot tell, the VBE's error is left unreported rather than
 * guessed. Reached through a function of the referenced project, the class's
 * members are not checked: XLIDE does not carry the function's return type
 * over. Nor are they through a default member, `ws(1).Nope`, which it does
 * not follow, nor is a member of an Enum variable, `c.Value` being "Invalid
 * qualifier". And these Word cases also reference Normal, before
 * ProbeLibrary, as a project XLIDE has not read: a bare `Thing`, `Widget` or
 * `Solo` could be Normal's. The engine reads Normal where it can; see the
 * tests below.
 */
const UNCHECKED = /_(function_chain_missing|default_member_missing|enum_variable_member)_compile$|^referenced_class_word_(private_bare_type|new_bare|predeclared_missing|enum_bare_missing|module_enum_bare_missing)_compile$/;

type References = ReturnType<typeof oracleProjectReferences>;

function errors(oracleCase: OracleCase, change: (references: References) => References = (references) => references): string[] {
	const modules = oracleCase.modules.map(oracleModuleInput);
	const project = buildVbaProjectIndex(modules, undefined, { conditionalCompilation: { projectConstants: {} } });
	const procedures = projectProcedureSignatures(project);
	const references = change(oracleProjectReferences(oracleCase));
	return modules.flatMap((mod) => analyzeVbaModuleSource({
		source: mod.source, moduleName: mod.moduleName, moduleKind: effectiveModuleKind(mod), host: oracleCase.host,
		...references,
		...projectAnalysisOptionsForModule(project, mod.moduleName, procedures),
	}).diagnostics.filter((diagnostic) => diagnostic.severity === 'error').map((diagnostic) => `${diagnostic.code}: ${diagnostic.message}`));
}

describe('the classes of a referenced VBA project', () => {
	it('has every measured case, in all four hosts', () => {
		expect(measured).toHaveLength(140);
		for (const host of [undefined, 'word', 'powerpoint', 'access']) {
			expect(measured.filter((oracleCase) => oracleCase.host === host)).toHaveLength(35);
		}
	});

	it.each(measured.filter((oracleCase) => !UNCHECKED.test(oracleCase.id)).map((oracleCase) => [oracleCase.id, oracleCase] as const))('%s agrees with the VBE', (_id, oracleCase) => {
		const found = errors(oracleCase);
		if (oracleCase.expected === 'accepted') {
			expect(found).toEqual([]);
		} else {
			expect(found).toHaveLength(1);
			// Each says what the VBE says, but `WithEvents ... As New`, which
			// the WithEvents rule words as it does for any class.
			const vbe = /Compile error: (.+?)\. Measured/.exec(oracleCase.description)?.[1];
			if (/_withevents_new_compile$/.test(oracleCase.id)) {
				expect(found[0]).toMatch(/^withevents-declaration: WithEvents variable 'mW' cannot be declared As New\./);
				return;
			}
			expect(found[0]).toContain(vbe === 'Variable not defined' || vbe === 'Method or data member not found' ? `${vbe}: ` : `VBE compile error: ${vbe}.`);
		}
	});

	it.each(measured.filter((oracleCase) => UNCHECKED.test(oracleCase.id)).map((oracleCase) => [oracleCase.id, oracleCase] as const))('%s is left unreported', (_id, oracleCase) => {
		expect(errors(oracleCase)).toEqual([]);
	});

});

// Two referenced projects, which the oracle does not make: the order of the
// references decides, as issue #1503 measured for a module and a procedure.
describe.each(['excel', 'word', 'powerpoint', 'access'])('which referenced project a bare name reaches in %s', (host) => {
	const first = referencedProjectSurface('First', [
		{ name: 'Palette', kind: 'class', source: 'Attribute VB_Exposed = True\r\nOption Explicit\r\nPublic Enum Solo\r\n    sOne = 1\r\nEnd Enum\r\n' },
	]);
	const second = referencedProjectSurface('Second', [
		{ name: 'Solo', kind: 'standard', source: 'Option Explicit\r\nPublic Function Good() As Long\r\nEnd Function\r\n' },
		{ name: 'Widget', kind: 'class', source: 'Attribute VB_Exposed = True\r\nOption Explicit\r\n' },
	]);
	const run = (code: string, libraries: string[], surfaces: ReferencedProjectSurface[] = [first, second]): string[] => analyzeVbaModuleSource({
		source: `Option Explicit\r\nPublic Sub Main()\r\n    Dim x As Variant\r\n    ${code}\r\nEnd Sub\r\n`,
		moduleName: 'Module1', moduleKind: 'standard', host, projectName: 'VBAProject',
		referencedLibraries: ['stdole', 'Office', ...libraries], referencedProjects: surfaces.map((surface) => surface.name), referencedProjectSurfaces: surfaces,
	}).diagnostics.filter((diagnostic) => diagnostic.severity === 'error').map((diagnostic) => diagnostic.message);

	it("is an earlier project's Enum before a later project's module of the name", () => {
		expect(run('x = Solo.sOne', ['First', 'Second'])).toEqual([]);
		expect(run('x = Solo.Nope', ['First', 'Second'])).toEqual(["Method or data member not found: 'Solo.Nope'."]);
	});

	it("is the module when its project comes first", () => {
		expect(run('x = Solo.Good()', ['Second', 'First'])).toEqual([]);
		expect(run('x = Solo.sOne', ['Second', 'First'])).toEqual(["Method or data member not found: 'Solo.sOne'."]);
	});

	it('is left unjudged where an unread library comes first and might have it', () => {
		expect(run('x = Solo.Nope', ['Unread', 'First'], [first])).toEqual([]);
		expect(run('Dim w As Object\r\n    Set w = New Widget', ['Unread', 'Second'], [second])).toEqual([]);
	});

	it("reports New on a class only as the name's first project offers it", () => {
		expect(run('Dim w As Object\r\n    Set w = New Widget', ['Second'], [second]).join()).toContain('Invalid use of New keyword');
	});
});

describe.each(['excel', 'word', 'powerpoint', 'access'])('a referenced project in %s', (host) => {
	const run = (modules: ReferencedProjectModuleSource[], code: string): string[] => {
		const source = `Option Explicit\r\nPublic Sub Main()\r\n    Dim x As Variant\r\n    ${code}\r\nEnd Sub\r\n`;
		return analyzeVbaModuleSource({
			source, moduleName: 'Module1', moduleKind: 'standard', host,
			referencedLibraries: ['stdole', 'Office', 'ProbeLibrary'], projectName: 'VBAProject',
			referencedProjects: ['ProbeLibrary'], referencedProjectSurfaces: [referencedProjectSurface('ProbeLibrary', modules)],
		}).diagnostics.filter((diagnostic) => diagnostic.severity === 'error').map((diagnostic) => diagnostic.message);
	};
	const enums: ReferencedProjectModuleSource[] = [
		{ name: 'Widget', kind: 'class', source: 'Attribute VB_Exposed = True\r\nOption Explicit\r\nPublic Enum WidgetColor\r\n    wcRed = 1\r\nEnd Enum\r\n' },
		{ name: 'Service', kind: 'standard', source: 'Option Explicit\r\nPublic Enum ServiceKind\r\n    skOne = 1\r\nEnd Enum\r\n' },
	];

	it("checks the member named after one of its Enums", () => {
		expect(run(enums, 'x = ProbeLibrary.WidgetColor.wcRed + ProbeLibrary.ServiceKind.skOne')).toEqual([]);
		expect(run(enums, 'x = ProbeLibrary.WidgetColor.Nope')).toEqual(["Method or data member not found: 'ProbeLibrary.WidgetColor.Nope'."]);
		expect(run(enums, 'x = ProbeLibrary.ServiceKind.Nope')).toEqual(["Method or data member not found: 'ProbeLibrary.ServiceKind.Nope'."]);
	});

	it('keeps an exposed document module a type and leaves its members unjudged', () => {
		// XLIDE keeps no member list for a document module: its host object's members are its own.
		const sheet: ReferencedProjectModuleSource[] = [{ name: 'Sheet1', kind: 'document', source: 'Attribute VB_Exposed = True\r\nOption Explicit\r\nPublic Sub Hello()\r\nEnd Sub\r\n' }];
		expect(run(sheet, 'Dim s As ProbeLibrary.Sheet1\r\n    ProbeLibrary.Sheet1.Hello\r\n    x = ProbeLibrary.Sheet1.Range("A1")')).toEqual([]);
		expect(referencedProjectSurface('ProbeLibrary', sheet).modules[0]).toEqual({ name: 'Sheet1', kind: 'document', visible: true, members: [], types: ['Sheet1'], enums: [] });
	});

	it("offers an exposed document module's Public Enums as an exposed class's", () => {
		const sheet: ReferencedProjectModuleSource[] = [{ name: 'Sheet1', kind: 'document', source: 'Attribute VB_Exposed = True\r\nOption Explicit\r\nPublic Enum SheetKind\r\n    shOne = 1\r\nEnd Enum\r\n' }];
		expect(run(sheet, 'Dim s As ProbeLibrary.SheetKind\r\n    x = ProbeLibrary.SheetKind.shOne + ProbeLibrary.shOne + shOne')).toEqual([]);
		expect(run(sheet, 'x = ProbeLibrary.SheetKind.Nope')).toEqual(["Method or data member not found: 'ProbeLibrary.SheetKind.Nope'."]);
		// Not exposed, it offers nothing.
		const hidden: ReferencedProjectModuleSource[] = [{ name: 'Sheet1', kind: 'document', source: sheet[0].source.replace('Attribute VB_Exposed = True\r\n', '') }];
		expect(referencedProjectSurface('ProbeLibrary', hidden).modules[0]).toEqual({ name: 'Sheet1', kind: 'document', visible: false, members: [], types: [] });
	});
});
