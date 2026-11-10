// stdole (OLE Automation), which every new Office file references, offers
// names without a qualifier: StdFunctions' LoadPicture and SavePicture, and
// the members of its OLE_TRISTATE and LoadPictureConstants enums. Measured in
// Excel, Word, PowerPoint and Access 16.0 (build 20527) through the VBE
// oracle: each compiles bare under Option Explicit, and so does
// `StdFunctions.LoadPicture("")`, while a bare `StdFunctions` does not, and
// SavePicture takes exactly two arguments. The cases are the stdole_name_*
// fixtures in syntax_corpus/oracle/vbe_oracle_cases.json.

import * as fs from 'node:fs';
import * as path from 'node:path';
import { describe, expect, it } from 'vitest';
import { analyzeModule } from '../../../src/analyzer';
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
) as { cases: OracleCase[] }).cases.filter((kase) => kase.id.startsWith('stdole_name_'));

/** Each error code, with the project known as the extension knows it, which turns on the undeclared-name check. */
function errorCodes(source: string, host = 'excel', references?: string[]): string[] {
	const modules = [oracleModuleInput({ name: 'Module1', type: 'standard', source })];
	const project = buildVbaProjectIndex(modules, undefined, { conditionalCompilation: { projectConstants: {} } });
	const procedures = projectProcedureSignatures(project);
	const known = oracleProjectReferences({ host });
	return modules.flatMap((mod) => analyzeVbaModuleSource({
		source: mod.source,
		moduleName: mod.moduleName,
		moduleKind: effectiveModuleKind(mod),
		host,
		...known,
		...(references ? { referencedLibraries: references } : {}),
		...projectAnalysisOptionsForModule(project, mod.moduleName, procedures),
	}).diagnostics.filter((diagnostic) => diagnostic.severity === 'error').map((diagnostic) => diagnostic.code ?? ''));
}

/** What XLIDE reports for each refused fixture. */
const REFUSED_AS: Readonly<Record<string, string>> = {
	stdfunctions_bare: 'undeclared-variable',
	stdfunctions_nope: 'undeclared-variable',
	savepicture_one: 'argument-count',
	savepicture_three: 'argument-count',
	savepicture_none: 'argument-count',
	savepicture_value: 'undeclared-variable',
	checked_assign: 'undeclared-variable',
	color_assign: 'undeclared-variable',
	for_gray: 'undeclared-variable',
	input_color: 'undeclared-variable',
	const_loadpicture: 'undeclared-variable',
};

/**
 * Refused by the VBE, not reported, and why. LoadPicture's arguments are
 * checked in a statement call only, not in an expression. Access's own
 * Application.LoadPicture takes one argument, and its arity is not checked;
 * as a member of Access's Application it was known on main too, so a Const's
 * value may use it there unreported.
 */
const KNOWN_QUIET: ReadonlyMap<string, { hosts: 'all' | readonly string[] }> = new Map([
	['loadpicture_five', { hosts: 'all' }],
	['loadpicture_four', { hosts: ['access'] }],
	['loadpicture_noarg', { hosts: ['access'] }],
	['const_loadpicture', { hosts: ['access'] }],
]);

const procedure = (...lines: string[]) => ['Option Explicit', 'Sub Main()', '    Dim p As Object, v As Variant', ...lines.map((line) => '    ' + line), 'End Sub', ''].join('\r\n');

describe("stdole's bare names against the measured fixtures", () => {
	it('has fixtures for every Office host', () => {
		expect(new Set(fixtures.map((fixture) => fixture.host))).toEqual(new Set(['excel', 'word', 'powerpoint', 'access']));
	});

	it.each(fixtures.map((fixture) => [fixture.id, fixture] as const))('%s', (id, fixture) => {
		const kind = id.replace(/^stdole_name_(excel|word|powerpoint|access)_/, '').replace(/_compile$/, '');
		const quiet = KNOWN_QUIET.get(kind);
		if (fixture.expected === 'accepted' || (quiet && (quiet.hosts === 'all' || quiet.hosts.includes(fixture.host)))) {
			expect(errorCodes(fixture.modules[0].source, fixture.host)).toEqual([]);
			return;
		}
		expect(errorCodes(fixture.modules[0].source, fixture.host)).toEqual([REFUSED_AS[kind]]);
	});
});

describe("stdole's bare names elsewhere", () => {
	it('are names in any case, read bare or called', () => {
		expect(errorCodes(procedure('Set p = loadpicture("")'))).toEqual([]);
		expect(errorCodes(procedure('Set p = LoadPicture'))).toEqual([]);
		expect(errorCodes(procedure('Call SavePicture(p, "x")'))).toEqual([]);
		expect(errorCodes(procedure('If v = GRAY Then v = vgacolor'))).toEqual([]);
	});

	it('stay known when the reference list leaves stdole out, which was not measured', () => {
		expect(errorCodes(procedure('Set p = LoadPicture("")'), 'excel', ['Office'])).toEqual([]);
	});

	it('read SavePicture, a Sub, as no value, and StdFunctions only before LoadPicture', () => {
		expect(errorCodes(procedure('v = SavePicture'))).toEqual(['undeclared-variable']);
		expect(errorCodes(procedure('Set p = StdFunctions _', '    .LoadPicture("")'))).toEqual([]);
		expect(errorCodes(procedure('Set p = StdFunctions.[LoadPicture]("")'))).toEqual([]);
		expect(errorCodes(procedure('v = StdFunctions.Nope'))).toEqual(['undeclared-variable']);
		// Its arguments would go unchecked, so a qualified SavePicture stays as main had it.
		expect(errorCodes(procedure('StdFunctions.SavePicture p, "x"'))).toEqual(['undeclared-variable']);
		expect(errorCodes(['Option Explicit', 'Const K = StdFunctions', ''].join('\r\n'))).toEqual(['undeclared-variable']);
	});

	it("leave a host's own LoadPicture its own arguments", () => {
		const vb6 = (...lines: string[]) => analyzeModule(procedure(...lines), { host: 'vb6', moduleName: 'Module1', moduleKind: 'standard' })
			.filter((diagnostic) => diagnostic.code === 'argument-count');
		expect(vb6('LoadPicture "", 0, 0, 1, 1')).toEqual([]);
		expect(errorCodes(procedure('Set p = LoadPicture("", 1, 1, 0)'), 'access')).toEqual([]);
	});

	it("check SavePicture's arguments against stdole's signature", () => {
		expect(errorCodes(procedure('SavePicture'))).toEqual(['argument-count']);
		expect(errorCodes(procedure('Call SavePicture()'))).toEqual(['argument-count']);
		expect(errorCodes(procedure('Call SavePicture(p, "x")'))).toEqual([]);
	});

	it('read a written name and a call through line continuations', () => {
		expect(errorCodes(procedure('Input #1, _', '    Color'))).toEqual(['undeclared-variable']);
		expect(errorCodes(procedure('For _', '    Gray = 1 To 3', 'Next'))).toEqual(['undeclared-variable']);
		expect(errorCodes(procedure('If v Then _', '    SavePicture p, "x"'))).toEqual([]);
	});

	it('read a write inside a one-line If, and LSet, RSet and Mid targets, as writes', () => {
		expect(errorCodes(procedure('If v Then Color = 1'))).toEqual(['undeclared-variable']);
		expect(errorCodes(procedure('If v Then Set Default = p'))).toEqual(['undeclared-variable']);
		expect(errorCodes(procedure('If v Then v = 1 Else Checked = 2'))).toEqual(['undeclared-variable']);
		expect(errorCodes(procedure('If v Then _', '    Color = 1'))).toEqual(['undeclared-variable']);
		expect(errorCodes(procedure('LSet Color = "x"'))).toEqual(['undeclared-variable']);
		expect(errorCodes(procedure('Mid(Color, 1, 1) = "x"'))).toEqual(['undeclared-variable']);
		expect(errorCodes(procedure('Mid$(Color, 1) = "x"'))).toEqual(['undeclared-variable']);
		expect(errorCodes(procedure('If v Then MidB$(Gray, 1) = "x"'))).toEqual(['undeclared-variable']);
		expect(errorCodes(procedure('If v = Color Then v = Gray'))).toEqual([]);
		expect(errorCodes(procedure('Color(1) = 1'))).toEqual(['undeclared-variable']);
		expect(errorCodes(procedure('If v Then Color(1) = 1'))).toEqual(['undeclared-variable']);
		expect(errorCodes(procedure('If v Then SavePicture = 1'))).toEqual(['undeclared-variable']);
		expect(errorCodes(procedure('If v Then v = 1 Else SavePicture = 1'))).toEqual(['undeclared-variable']);
		expect(errorCodes(procedure('Color(v(1)) = 1'))).toEqual(['undeclared-variable']);
		expect(errorCodes(procedure('If v Then Checked'))).toEqual(['undeclared-variable']);
		expect(errorCodes(procedure('If v Then Color 1, 2'))).toEqual(['undeclared-variable']);
		expect(errorCodes(procedure('If v Then Set p = LoadPicture("") Else LoadPicture ""'))).toEqual([]);
		expect(errorCodes(procedure('If v Then Call Checked'))).toEqual(['undeclared-variable']);
		expect(errorCodes(procedure('Color(v(v(v(1)))) = 1'))).toEqual(['undeclared-variable']);
		expect(errorCodes(procedure('Color(v("(")) = 1'))).toEqual(['undeclared-variable']);
		// A Single's type suffix and then Then or Else, which aren't member names.
		expect(errorCodes(procedure('Dim x As Single', 'If x! Then Color = 1'))).toEqual(['undeclared-variable']);
		expect(errorCodes(procedure('Dim x As Single', 'If True Then x = x! Else Color = 1'))).toEqual(['undeclared-variable']);
		expect(errorCodes(procedure('Dim x As Single', 'If x! Then SavePicture p, "x"'))).toEqual([]);
	});

	it("read a file statement's file and record numbers as values, and its variables as written", () => {
		expect(errorCodes(procedure('Get #1, Checked, v'))).toEqual([]);
		expect(errorCodes(procedure('Input #Checked, v'))).toEqual([]);
		expect(errorCodes(procedure('Put #1, , Checked'))).toEqual([]);
		expect(errorCodes(procedure('Get #1, , Checked'))).toEqual(['undeclared-variable']);
		expect(errorCodes(procedure('Input #1, v, Color'))).toEqual(['undeclared-variable']);
		expect(errorCodes(procedure('Line Input #1, Gray'))).toEqual(['undeclared-variable']);
		expect(errorCodes(procedure('Input #1, v(Checked)'))).toEqual([]);
		expect(errorCodes(procedure('Input #1, v(":"), Color'))).toEqual(['undeclared-variable']);
		expect(errorCodes(procedure('Input #1, v.Then, Color'))).toEqual(['undeclared-variable']);
		expect(errorCodes(procedure('Get #1, v!Else, Gray'))).toEqual(['undeclared-variable']);
	});

	it('are only those names', () => {
		expect(errorCodes(procedure('Set p = LoadPictures("")'))).toEqual(['undeclared-variable']);
		expect(errorCodes(procedure('v = Grey'))).toEqual(['undeclared-variable']);
	});
});
