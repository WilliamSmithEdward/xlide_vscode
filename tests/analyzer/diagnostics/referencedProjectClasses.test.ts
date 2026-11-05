// The classes of a referenced VBA project, measured through the VBE oracle in
// Excel, Word, PowerPoint and Access 16.0 (build 20527), every host alike:
// ProbeLibrary's class Widget is exposed (VB_Exposed, Instancing
// PublicNotCreatable), Solo is exposed with a default instance and Thing is
// private. The cases are the referenced_class_* fixtures in
// syntax_corpus/oracle/vbe_oracle_cases.json.

import * as fs from 'node:fs';
import * as path from 'node:path';
import { describe, expect, it } from 'vitest';
import { referencedProjectSurface, type ReferencedProjectModule, type ReferencedProjectModuleSource, type ReferencedProjectSurface } from '../../../src/analyzer/symbols/referencedProjects';
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
 * over. Nor is a member of an Enum variable, `c.Value` being "Invalid
 * qualifier". A default member is followed where it takes the index and
 * returns an exposed class, `ws(1).Nope` (issue #1623). And these Word cases also reference Normal, before
 * ProbeLibrary, as a project XLIDE has not read: a bare `Thing`, `Widget` or
 * `Solo` could be Normal's. The engine reads Normal where it can; see the
 * tests below.
 */
const UNCHECKED = /_(function_chain_missing|enum_variable_member)_compile$|^referenced_class_word_(private_bare_type|new_bare|predeclared_missing|enum_bare_missing|module_enum_bare_missing|bare_missing_member)_compile$/;

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
		expect(measured).toHaveLength(224);
		for (const host of [undefined, 'word', 'powerpoint', 'access']) {
			expect(measured.filter((oracleCase) => oracleCase.host === host)).toHaveLength(56);
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

// What the oracle's one provider cannot hold (issue #1603): a second project,
// a member named Item, a run-time function's name and a continued member type.
describe.each(['excel', 'word', 'powerpoint', 'access'])('the members of a referenced class in %s', (host) => {
	const exposed = 'Attribute VB_Exposed = True\r\nOption Explicit\r\n';
	const predeclared = `Attribute VB_PredeclaredId = True\r\n${exposed}`;
	const run = (code: string, projects: Record<string, ReferencedProjectModuleSource[]>, params = ''): string[] => {
		const surfaces = Object.entries(projects).map(([name, modules]) => referencedProjectSurface(name, modules));
		return analyzeVbaModuleSource({
			source: `Option Explicit\r\nPublic Sub Main(${params})\r\n    Dim x As Variant\r\n    ${code}\r\nEnd Sub\r\n`,
			moduleName: 'Module1', moduleKind: 'standard', host, projectName: 'VBAProject',
			referencedLibraries: ['stdole', 'Office', ...Object.keys(projects)], referencedProjects: Object.keys(projects), referencedProjectSurfaces: surfaces,
		}).diagnostics.filter((diagnostic) => diagnostic.severity === 'error').map((diagnostic) => diagnostic.message);
	};
	const solo: ReferencedProjectModuleSource = { name: 'Solo', kind: 'class', source: `${predeclared}Public Value As Long\r\n` };

	it("takes an earlier project's Enum of the name before a later default instance", () => {
		const palette: ReferencedProjectModuleSource = { name: 'Palette', kind: 'class', source: `${exposed}Public Enum Solo\r\n    sOne = 1\r\nEnd Enum\r\n` };
		expect(run('x = Solo.sOne', { First: [palette], Second: [solo] })).toEqual([]);
	});

	it("lets no earlier project's hidden module stop a later default instance", () => {
		const hidden: ReferencedProjectModuleSource = { name: 'Solo', kind: 'class', source: 'Option Explicit\r\n' };
		const secret: ReferencedProjectModuleSource = { name: 'Solo', kind: 'standard', source: 'Option Private Module\r\nOption Explicit\r\n' };
		for (const earlier of [hidden, secret]) {
			expect(run('x = Solo.Value', { First: [earlier], Second: [solo] })).toEqual([]);
			expect(run('x = Solo.Nope', { First: [earlier], Second: [solo] }).join()).toContain('Nope');
		}
	});

	it('leaves a run-time function ahead of a default instance of its name', () => {
		const shadow: ReferencedProjectModuleSource = { name: 'CreateObject', kind: 'class', source: `${predeclared}Public Value As Long\r\n` };
		expect(run('x = CreateObject("Excel.Application").Visible', { ProbeLibrary: [shadow] })).toEqual([]);
	});

	it("indexes what a parameterless Item returns, whatever it is called", () => {
		const modules: ReferencedProjectModuleSource[] = [
			{ name: 'Holder', kind: 'class', source: `${exposed}Public Function Item() As Widgets\r\nEnd Function\r\nPublic Function Make(ByVal n As Long) As Widgets\r\nEnd Function\r\n` },
			{ name: 'Widgets', kind: 'class', source: `${exposed}Public Own As Long\r\nPublic Function Get1(ByVal i As Long) As Object\r\nAttribute Get1.VB_UserMemId = 0\r\nEnd Function\r\n` },
		];
		const params = 'ByVal h As ProbeLibrary.Holder';
		// `Item(1)` indexes Widgets, so what it gives is unjudged. `Make(1)` is Make's own argument.
		expect(run('x = h.Item(1).Anything', { ProbeLibrary: modules }, params)).toEqual([]);
		expect(run('x = h.Make(1).Own', { ProbeLibrary: modules }, params)).toEqual([]);
		expect(run('x = h.Make(1).Nope', { ProbeLibrary: modules }, params).join()).toContain('Nope');
	});

	it.each(['\n', '\r\n', '\r'])('reads a member type continued with %j lines', (eol) => {
		const modules: ReferencedProjectModuleSource[] = [{ name: 'Widget', kind: 'class', source: `${exposed}Public Child As ProbeLibrary. _${eol}    Widget\r\nPublic Value As Long\r\n` }];
		const params = 'ByVal w As ProbeLibrary.Widget';
		expect(run('x = w.Child.Value', { ProbeLibrary: modules }, params)).toEqual([]);
		expect(run('x = w.Child.Nope', { ProbeLibrary: modules }, params).join()).toContain('Nope');
	});

	it('reads the module list a bounded number of times, however many names refer to it', () => {
		const modules: ReferencedProjectModule[] = Array.from({ length: 300 }, (_, i) => ({
			name: `Part${i}`, kind: 'class', visible: true, types: [`Part${i}`], members: [], objectMembers: [{ name: 'Value', kind: 'property', takesArguments: false }],
		}));
		let reads = 0;
		const surface: ReferencedProjectSurface = {
			name: 'ProbeLibrary',
			modules: new Proxy(modules, {
				get(target, key, receiver) {
					if (typeof key === 'string' && /^\d+$/.test(key)) { reads++; }
					return Reflect.get(target, key, receiver);
				},
			}),
		};
		const params = modules.slice(0, 25).map((module, i) => `ByVal p${i} As ProbeLibrary.${module.name}`).join(', ');
		const body = modules.slice(0, 25).map((_, i) => `x = p${i}.Value`).join('\r\n    ');
		const found = analyzeVbaModuleSource({
			source: `Option Explicit\r\nPublic Sub Main(${params})\r\n    Dim x As Variant\r\n    ${body}\r\nEnd Sub\r\n`,
			moduleName: 'Module1', moduleKind: 'standard', host, projectName: 'VBAProject',
			referencedLibraries: ['stdole', 'Office', 'ProbeLibrary'], referencedProjects: ['ProbeLibrary'], referencedProjectSurfaces: [surface],
		}).diagnostics.filter((diagnostic) => diagnostic.severity === 'error');
		expect(found).toEqual([]);
		expect(reads).toBeLessThan(modules.length * 4);
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
