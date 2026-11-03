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
 * ProbeLibrary, as a project XLIDE has not read: a bare `Thing`, `Widget`,
 * `Solo`, `WidgetColor` or `ServiceKind` could be Normal's. The engine
 * reads Normal where it can; see the tests below.
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
		expect(measured).toHaveLength(232);
		for (const host of [undefined, 'word', 'powerpoint', 'access']) {
			expect(measured.filter((oracleCase) => oracleCase.host === host)).toHaveLength(58);
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

	it('names what the class lacks', () => {
		expect(errors(byId('referenced_class_missing_member_compile'))).toEqual(["member-not-found: Method or data member not found: 'ProbeLibrary.Widget.Nope'."]);
		expect(errors(byId('referenced_class_friend_member_compile'))).toEqual(["member-not-found: Method or data member not found: 'ProbeLibrary.Widget.Inner'."]);
		expect(errors(byId('referenced_class_class_as_value_compile'))).toEqual(["member-not-found: Method or data member not found: 'ProbeLibrary.Widget'."]);
		expect(errors(byId('referenced_class_new_bare_compile'))).toEqual(["invalid-new-type-name: 'Widget' is a class of the VBA project ProbeLibrary, which another project cannot create. This is a VBE compile error: Invalid use of New keyword."]);
	});
});

describe('what an exposed class of a referenced project offers', () => {
	const surface = referencedProjectSurface('ProbeLibrary', byId('referenced_class_qualified_type_compile').projectReference.modules.map((mod) => ({
		name: mod.name, kind: mod.type === 'class' ? 'class' as const : 'standard' as const, source: mod.source,
	})));
	const module = (name: string) => surface.modules.find((candidate) => candidate.name === name);

	it('is its Public members and Enums, and no Private, Friend or event member', () => {
		expect(module('Widget')).toMatchObject({
			visible: true,
			types: ['Widget', 'WidgetColor'],
			enums: [{ name: 'WidgetColor', members: ['wcRed'] }],
			predeclared: false,
			defaultMember: false,
		});
		expect(module('Widgets')).toMatchObject({ visible: true, defaultMember: true });
		expect(module('Service')).toMatchObject({ types: ['ServiceKind'], enums: [{ name: 'ServiceKind', members: ['skOne'] }] });
		expect(module('Widget')?.objectMembers?.map((member) => member.name)).toEqual(['Value', 'Child', 'Size', 'Twice', 'Poke']);
		expect(module('Solo')).toMatchObject({ visible: true, predeclared: true });
	});

	it('is nothing for a private class, its Enum included', () => {
		expect(module('Thing')).toEqual({ name: 'Thing', kind: 'class', visible: false, members: [], types: [] });
	});
});

describe('which referenced project a bare name reaches', () => {
	/** A case's own code, its references as the oracle's Excel file has them. */
	const withCode = (code: string, extra: OracleFixtureModule[] = []): OracleCase => ({
		...byId('referenced_class_qualified_type_compile'),
		modules: [{ name: 'Module1', type: 'standard', source: `Option Explicit\r\nPublic Sub Main()\r\n    Dim x As Long, v As Variant\r\n    ${code.replace(/\n/g, '\r\n')}\r\nEnd Sub\r\n` }, ...extra],
	});
	const project = (name: string, modules: OracleFixtureModule[]): ReferencedProjectSurface => referencedProjectSurface(name, modules.map((mod) => ({
		name: mod.name, kind: mod.type === 'class' ? 'class' as const : 'standard' as const, source: mod.source,
	})));
	const before = (surface: ReferencedProjectSurface) => (references: References): References => ({
		...references,
		referencedLibraries: ['stdole', 'Office', surface.name, 'ProbeLibrary'],
		referencedProjectSurfaces: [surface, ...references.referencedProjectSurfaces!],
	});

	it('is the first project that has it, a module before a later class', () => {
		// First's standard module Solo wins over ProbeLibrary's predeclared class Solo.
		const first = project('First', [{ name: 'Solo', type: 'standard', source: 'Option Explicit\r\nPublic Function Good() As Long\r\nEnd Function\r\n' }]);
		expect(errors(withCode('x = Solo.Good()'), before(first))).toEqual([]);
		expect(errors(withCode('x = Solo.Good()'))).toEqual(["member-not-found: Method or data member not found: 'ProbeLibrary.Solo.Good'."]);
	});

	it("is the first project's type, an earlier project's class winning", () => {
		const first = project('First', [{ name: 'Widget', type: 'class', source: 'Attribute VB_Exposed = True\r\nOption Explicit\r\nPublic Mine As Long\r\n' }]);
		expect(errors(withCode('Dim w As Widget\n    x = w.Mine'), before(first))).toEqual([]);
		expect(errors(withCode('Dim w As Widget\n    x = w.Value'), before(first))).toEqual(["member-not-found: Method or data member not found: 'First.Widget.Value'."]);
		// Qualified, it is the project named.
		expect(errors(withCode('Dim w As ProbeLibrary.Widget\n    x = w.Mine'), before(first))).toEqual(["member-not-found: Method or data member not found: 'ProbeLibrary.Widget.Mine'."]);
	});

	it('is none where a library before the project might have it', () => {
		const unread = (references: References): References => ({ ...references, referencedLibraries: ['stdole', 'Unread', 'Office', 'ProbeLibrary'] });
		expect(errors(withCode('Dim w As Widget\n    Set w = New Widget'), unread)).toEqual([]);
		expect(errors(withCode('x = Solo.Nope'), unread)).toEqual([]);
		expect(errors(withCode('Dim w As Widget\n    x = w.Nope'), unread)).toEqual([]);
		// Qualified, the name is the project's whatever comes before it.
		expect(errors(withCode('Dim w As ProbeLibrary.Widget\n    Set w = New ProbeLibrary.Widget\n    x = w.Nope'), unread)).toHaveLength(2);
	});

	it("is ProbeLibrary's in Word once Normal is read and lacks it", () => {
		const normal = project('Normal', [{ name: 'NewMacros', type: 'standard', source: 'Option Explicit\r\n' }]);
		const read = (references: References): References => ({ ...references, referencedProjectSurfaces: [normal, ...references.referencedProjectSurfaces!] });
		for (const id of ['referenced_class_word_private_bare_type_compile', 'referenced_class_word_new_bare_compile', 'referenced_class_word_predeclared_missing_compile',
			'referenced_class_word_enum_bare_missing_compile', 'referenced_class_word_module_enum_bare_missing_compile']) {
			expect(errors(byId(id), read), id).toHaveLength(1);
		}
	});
});

describe('how a referenced class is named', () => {
	const withCode = (code: string): OracleCase => ({
		...byId('referenced_class_qualified_type_compile'),
		modules: [{ name: 'Module1', type: 'standard', source: `Option Explicit\r\nPublic Sub Main()\r\n    Dim x As Long\r\n    ${code.replace(/\n/g, '\r\n')}\r\nEnd Sub\r\n` }],
	});

	it('across a line continuation', () => {
		expect(errors(withCode('Dim w As ProbeLibrary. _\n        Widget\n    x = w.Value'))).toEqual([]);
		expect(errors(withCode('Dim w As ProbeLibrary. _\n        Widget\n    x = w.Nope'))).toEqual(["member-not-found: Method or data member not found: 'ProbeLibrary.Widget.Nope'."]);
		expect(errors(withCode('Dim t As ProbeLibrary. _\n        Thing'))[0]).toMatch(/offers no type 'Thing'/);
	});

	it('with a variable as its default member', () => {
		const surface = referencedProjectSurface('ProbeLibrary', [
			{ name: 'Bag', kind: 'class', source: 'Attribute VB_Exposed = True\r\nOption Explicit\r\nPublic Items As Collection\r\nAttribute Items.VB_VarUserMemId = 0\r\n' },
		]);
		expect(surface.modules[0]).toMatchObject({ visible: true, defaultMember: true });
	});

	it('by a member of the library itself, qualified with its own project name', () => {
		const surface = referencedProjectSurface('ProbeLibrary', [
			{ name: 'Node', kind: 'class', source: 'Attribute VB_Exposed = True\r\nOption Explicit\r\nPublic Parent As ProbeLibrary.Node\r\nPublic Other As OtherLib.Node\r\nPublic Function Up() As Node\r\nEnd Function\r\n' },
		]);
		expect(surface.modules[0].objectMembers).toEqual([
			{ name: 'Parent', kind: 'property', type: 'Node' },
			{ name: 'Other', kind: 'property' },
			{ name: 'Up', kind: 'method', type: 'Node' },
		]);
	});
});

describe('the work a large referenced project costs', () => {
	it('reads its module list a bounded number of times, however many names refer to it', () => {
		const classes = Array.from({ length: 300 }, (_, i): OracleFixtureModule => ({ name: `Part${i}`, type: 'class', source: 'Attribute VB_Exposed = True\r\nOption Explicit\r\nPublic Value As Long\r\n' }));
		const surface = referencedProjectSurface('ProbeLibrary', classes.map((mod) => ({ name: mod.name, kind: 'class' as const, source: mod.source })));
		let reads = 0;
		const counted: ReferencedProjectSurface = {
			name: surface.name,
			modules: new Proxy(surface.modules, {
				get(target, key, receiver) {
					if (typeof key === 'string' && /^\d+$/.test(key)) {
						reads += 1;
					}
					return Reflect.get(target, key, receiver);
				},
			}),
		};
		const lines = classes.map((mod, i) => `    Dim a${i} As ProbeLibrary.${mod.name}, b${i} As ${mod.name}\r\n    x = a${i}.Value + b${i}.Value`);
		const source = `Option Explicit\r\nPublic Sub Main()\r\n    Dim x As Long\r\n${lines.join('\r\n')}\r\nEnd Sub\r\n`;
		const found = analyzeVbaModuleSource({
			source, moduleName: 'Module1', moduleKind: 'standard', host: 'excel',
			referencedLibraries: ['stdole', 'Office', 'ProbeLibrary'], projectName: 'VBAProject',
			referencedProjects: ['ProbeLibrary'], referencedProjectSurfaces: [counted],
		}).diagnostics.filter((diagnostic) => diagnostic.severity === 'error');
		expect(found).toEqual([]);
		// Each rule that indexes the project reads it once. Building its
		// type list anew for each of the 300 qualified type names, as
		// missingReference.ts did before, read it 90,600 times.
		expect(reads).toBeLessThan(10 * classes.length);
	});
});

describe.each(['excel', 'word', 'powerpoint', 'access'])('a referenced project in %s', (host) => {
	const run = (modules: ReferencedProjectModuleSource[], code: string): string[] => {
		const source = `Option Explicit\r\nPublic Function F() As Object\r\nEnd Function\r\nPublic Function G(ByVal b As Boolean) As Object\r\nEnd Function\r\nPublic Sub Main()\r\n    Dim x As Variant\r\n    ${code}\r\nEnd Sub\r\n`;
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

	it('checks the member named after an Enum it reaches by bare name', () => {
		expect(run(enums, 'x = WidgetColor.wcRed + ServiceKind.skOne')).toEqual([]);
		expect(run(enums, 'x = WidgetColor.Nope')).toEqual(["Method or data member not found: 'WidgetColor.Nope'."]);
		expect(run(enums, 'x = ServiceKind.Nope')).toEqual(["Method or data member not found: 'ServiceKind.Nope'."]);
		// An exposed document module's Enum, bare.
		const sheet: ReferencedProjectModuleSource[] = [{ name: 'Sheet1', kind: 'document', source: 'Attribute VB_Exposed = True\r\nOption Explicit\r\nPublic Enum SheetKind\r\n    shOne = 1\r\nEnd Enum\r\n' }];
		expect(run(sheet, 'Dim s As SheetKind\r\n    x = SheetKind.shOne')).toEqual([]);
		expect(run(sheet, 'x = SheetKind.Nope')).toEqual(["Method or data member not found: 'SheetKind.Nope'."]);
		// The project's own name comes first.
		expect(run(enums, 'Dim WidgetColor As New Collection\r\n    x = WidgetColor.Count')).toEqual([]);
	});

	it("leaves an Enum named like a value VBA or the host gives first", () => {
		const global = { excel: 'ActiveWorkbook', word: 'ActiveDocument', powerpoint: 'ActivePresentation', access: 'CurrentProject' }[host]!;
		const named = (name: string): ReferencedProjectModuleSource[] => [{ name: 'Service', kind: 'standard', source: `Option Explicit\r\nPublic Enum ${name}\r\n    zOne = 1\r\nEnd Enum\r\n` }];
		expect(run(named(global), `x = ${global}.Name`)).toEqual([]);
		expect(run(named('Strings'), 'x = Strings.Left("ab", 1)')).toEqual([]);
		expect(run(named('Err'), 'x = Err.Number')).toEqual([]);
	});

	const widget: ReferencedProjectModuleSource[] = [
		{ name: 'Widget', kind: 'class', source: 'Attribute VB_Exposed = True\r\nOption Explicit\r\nPublic Value As Long\r\n' },
		{ name: 'Solo', kind: 'class', source: 'Attribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = True\r\nOption Explicit\r\nPublic Value As Long\r\n' },
	];

	it('takes a class after TypeOf ... Is as a type whatever the operand', () => {
		const typeOf = (operand: string) => `Dim c As New Collection, o As Object\r\n    If TypeOf ${operand} Is ProbeLibrary.Widget Then x = 1`;
		for (const operand of ['o', 'F()', 'c(1)', 'c.Item(1)', 'G(TypeOf o Is ProbeLibrary.Widget)', 'G(o Is Nothing)', 'o!Item', 'o![Some Item].Value', 'c(1)!Item']) {
			expect(run(widget, `${typeOf(operand)}\r\n    x = TypeOf ${operand} Is ProbeLibrary.Widget`).filter((message) => /ProbeLibrary/.test(message)), operand).toEqual([]);
		}
		// On a Do block's closing line too.
		for (const loop of ['Until', 'While']) {
			expect(run(widget, `Dim o As Object\r\n    Do\r\n        x = 1\r\n    Loop ${loop} TypeOf o Is ProbeLibrary.Widget`), loop).toEqual([]);
		}
	});

	it('reports a class with no default instance as a value in a block header too', () => {
		for (const header of ['If ProbeLibrary.Widget Is Nothing Then\r\n    End If', 'With ProbeLibrary.Widget\r\n    End With', 'Do While ProbeLibrary.Widget Is Nothing\r\n    Loop', 'Do\r\n        x = 1\r\n    Loop Until ProbeLibrary.Widget Is Nothing']) {
			expect(run(widget, header), header).toEqual(["Method or data member not found: 'ProbeLibrary.Widget'."]);
		}
		expect(run(widget, 'With ProbeLibrary.Solo\r\n    End With')).toEqual([]);
		expect(run(widget, 'If ProbeLibrary.Nope Then\r\n    End If')).toEqual(["Method or data member not found: 'ProbeLibrary.Nope'."]);
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

// Synthetic safety checks supplement the historical four-host fixtures.
describe.each(['excel', 'word', 'powerpoint', 'access'])('referenced type ordering in %s', host => {
    const run = (libraries: string[], surfaces: ReferencedProjectSurface[], code = 'Dim w As Widget: Set w = New Widget') => errors({
        ...byId('referenced_class_qualified_type_compile'), host,
        modules: [{ name: 'Module1', type: 'standard', source: `Option Explicit\r\nSub Main()\r\n${code}\r\nEnd Sub` }],
    }, refs => ({ ...refs, referencedLibraries: libraries, referencedProjectSurfaces: surfaces }));
    const provider = referencedProjectSurface('ProbeLibrary', [{ name: 'Widget', kind: 'class', source: 'Attribute VB_Exposed = True\r\nPublic Value As Long' }]);
    it('does not infer a bare class through an unread earlier reference', () => {
        expect(run(['Unread', 'ProbeLibrary'], [provider])).toEqual([]);
        expect(run(['Unread', 'ProbeLibrary'], [provider], 'Dim w As ProbeLibrary.Widget: Set w = New ProbeLibrary.Widget')).toHaveLength(1);
    });
    it('uses the earlier read project type', () => {
        const first = referencedProjectSurface('First', [{ name: 'Types', kind: 'standard', source: 'Public Enum Widget\r\nOne = 1\r\nEnd Enum' }]);
        expect(run(['First', 'ProbeLibrary'], [first, provider], 'Dim w As Widget')).toEqual([]);
        expect(run(['ProbeLibrary'], [provider])).toHaveLength(1);
    });
});
