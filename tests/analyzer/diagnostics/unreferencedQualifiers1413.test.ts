// Issue #1413, measured in Excel, Word, PowerPoint and Access 16.0 (build
// 20430) through the VBE oracle: a type qualified by a name that is no
// library the project references and none of its modules does not compile,
// "User-defined type not defined", after As, New and Implements, in a
// Declare too. The dir stream records each reference by the name its type
// library gives itself, and leaves out VBA and the host's own library. The
// cases are in syntax_corpus/oracle/vbe_oracle_cases.json.

import * as fs from 'node:fs';
import * as path from 'node:path';
import { describe, expect, it } from 'vitest';
import { hostTokenForLibid } from '../../../src/analyzer/host/hostLibraries';
import { accessVbaCfb } from '../../../src/vba/accessDatabase';
import { openMacroContainer } from '../../../src/vba/macroContainer';
import { VbaProject } from '../../../src/vba/vbaProject';
import { analyzeVbaModuleSource } from '../../../src/vbaModuleAnalysis';
import {
	buildVbaProjectIndex,
	effectiveModuleKind,
	projectAnalysisOptionsForModule,
	projectProcedureSignatures,
} from '../../../src/vbaProjectAnalysis';
import { oracleModuleInput, type OracleFixtureModule } from '../../helpers/oracleModuleInput';

type Host = 'excel' | 'word' | 'powerpoint' | 'access';
interface OracleCase {
	id: string;
	expected: 'accepted' | 'rejected';
	host?: Host;
	references?: { guid: string }[];
	modules: OracleFixtureModule[];
}

const corpus: { cases: OracleCase[] } = JSON.parse(
	fs.readFileSync(path.join(process.cwd(), 'syntax_corpus', 'oracle', 'vbe_oracle_cases.json'), 'utf8'),
);
const HOSTS = ['', 'word_', 'powerpoint_', 'access_'] as const;
/** The names the cases' references go by, as the dir stream records them (measured). */
const LIBRARY_NAMES: Readonly<Record<string, string>> = {
	'{420B2830-E718-11CF-893D-00A0C9054228}': 'Scripting',
	'{3F4DACA7-160D-11D2-A8E9-00104B365C9F}': 'VBScript_RegExp_55',
	'{B691E011-1797-432E-907A-4D8C69339129}': 'ADODB',
	'{F5078F18-C551-11D3-89B9-0000F81FE221}': 'MSXML2',
	'{0002E157-0000-0000-C000-000000000046}': 'VBIDE',
	'{662901FC-6951-4854-9EB2-D9A2570F2B2E}': 'WinHttp',
	'{50A7E9B0-70EF-11D1-B75A-00A0C90564FE}': 'Shell32',
	'{00062FFF-0000-0000-C000-000000000046}': 'Outlook',
};
/** What a new file of each host records before any reference is added, and its project's name (measured). */
const DEFAULTS: Readonly<Record<Host, readonly string[]>> = {
	excel: ['stdole', 'Office'], word: ['stdole', 'Normal', 'Office'], powerpoint: ['stdole', 'Office'], access: ['stdole', 'DAO'],
};
const PROJECT_NAMES: Readonly<Record<Host, string>> = { excel: 'VBAProject', word: 'Project', powerpoint: 'VBAProject', access: 'Database1' };

function caseFor(id: string): OracleCase {
	const found = corpus.cases.find((oracleCase) => oracleCase.id === id);
	if (!found) {
		throw new Error(`no oracle case ${id}`);
	}
	return found;
}

/** Each error code, with the project's references and name known, as the extension has them. */
function errorCodes(oracleCase: OracleCase, known = true): string[] {
	const host = oracleCase.host ?? 'excel';
	const modules = oracleCase.modules.map(oracleModuleInput);
	const project = buildVbaProjectIndex(modules, undefined, { conditionalCompilation: { projectConstants: {} } });
	const procedures = projectProcedureSignatures(project);
	const references = (oracleCase.references ?? []).map((one) => one.guid);
	return modules.flatMap((mod) => analyzeVbaModuleSource({
		source: mod.source,
		moduleName: mod.moduleName,
		moduleKind: effectiveModuleKind(mod),
		host,
		referencedHosts: references.flatMap((guid) => hostTokenForLibid(guid) ?? []),
		...(known ? { referencedLibraries: [...DEFAULTS[host], ...references.map((guid) => LIBRARY_NAMES[guid])], projectName: PROJECT_NAMES[host] } : {}),
		...projectAnalysisOptionsForModule(project, mod.moduleName, procedures),
	}).diagnostics.filter((diagnostic) => diagnostic.severity === 'error').map((diagnostic) => diagnostic.code ?? ''));
}

const eachHost = (names: readonly string[]): Array<[string]> =>
	names.flatMap((name) => HOSTS.map((host): [string] => [`issue1413_${host}${name}_compile`]));
const LIBRARIES = ['scripting', 'regexp', 'adodb', 'msxml', 'vbide', 'winhttp', 'shell', 'outlook'];

describe('a type qualified by a library the project does not reference (#1413)', () => {
	it.each(eachHost([
		...LIBRARIES.map((name) => `noref_${name}`), 'unknown_qualifier', 'unknown_new', 'unknown_module_level',
		'unknown_parameter', 'unknown_implements', 'unknown_declare',
	]))('%s is User-defined type not defined', (id) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('rejected');
		expect(errorCodes(oracleCase)).toEqual(['missing-library-reference']);
	});

	it.each(eachHost([
		...LIBRARIES.map((name) => `ref_${name}`), 'ref_lowercase_qualifier', 'ref_declare_outlook', 'default_vba',
		'default_stdole', 'module_type', 'other_module_class', 'string_holds_qualifier',
	]))('%s compiles, and nothing reports it', (id) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('accepted');
		expect(errorCodes(oracleCase)).toEqual([]);
	});

	it.each(eachHost(['noref_adodb', 'unknown_qualifier', 'unknown_implements']))('%s waits for the references to be known', (id) => {
		expect(errorCodes(caseFor(id), false)).toEqual([]);
	});

	it.each([
		['issue1413_default_office_compile', 'accepted'],
		['issue1413_word_default_office_compile', 'accepted'],
		['issue1413_powerpoint_default_office_compile', 'accepted'],
		['issue1413_access_default_office_compile', 'rejected'],
	] as const)('%s: Office is a default reference everywhere but Access (%s)', (id, expected) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe(expected);
		expect(errorCodes(oracleCase)).toEqual(expected === 'accepted' ? [] : ['missing-library-reference']);
	});

	it.each([
		['issue1413_project_vbaproject_class_compile', 'accepted'],
		['issue1413_word_project_vbaproject_class_compile', 'rejected'],
		['issue1413_powerpoint_project_vbaproject_class_compile', 'accepted'],
		['issue1413_access_project_vbaproject_class_compile', 'rejected'],
		['issue1413_project_project_class_compile', 'rejected'],
		['issue1413_word_project_project_class_compile', 'accepted'],
		['issue1413_powerpoint_project_project_class_compile', 'rejected'],
		['issue1413_access_project_project_class_compile', 'rejected'],
	] as const)("%s: the project's own name qualifies its types, and only it (%s)", (id, expected) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe(expected);
		expect(errorCodes(oracleCase)).toEqual(expected === 'accepted' ? [] : ['missing-library-reference']);
	});

	it('names the qualifier, offers no fix it cannot carry out, and reports a library once per module', () => {
		const source = 'Option Explicit\r\nPrivate a As Nowhere.Thing\r\nPrivate b As Nowhere.Other\r\nPrivate c As Elsewhere.Thing\r\n';
		const project = buildVbaProjectIndex([{ moduleName: 'Module1', source }], undefined, { conditionalCompilation: { projectConstants: {} } });
		const found = analyzeVbaModuleSource({
			source, moduleName: 'Module1', moduleKind: 'standard', referencedLibraries: ['stdole', 'Office'], projectName: 'VBAProject',
			...projectAnalysisOptionsForModule(project, 'Module1', projectProcedureSignatures(project)),
		}).diagnostics.filter((diagnostic) => diagnostic.code === 'missing-library-reference');
		expect(found.map((diagnostic) => source.slice(diagnostic.span.start, diagnostic.span.end))).toEqual(['Nowhere.Thing', 'Elsewhere.Thing']);
		expect(found[0].message).toContain("'Nowhere' is no library this project references and none of its modules");
		expect(found[0].message).toContain('User-defined type not defined');
		expect(found.every((diagnostic) => diagnostic.data?.addLibraryReference === undefined)).toBe(true);
	});

	it('leaves Scripting to its own check, so a missing Scripting reference is reported once', () => {
		expect(errorCodes(caseFor('issue1413_noref_scripting_compile'))).toEqual(['missing-library-reference']);
	});
});

describe("the project's own name, from the dir stream", () => {
	const binaries = path.join(__dirname, '..', '..', 'fixtures', 'binaries');
	it.each([
		['FormFixture.xlsm', 'VBAProject'],
		['WordFixture.docm', 'Project'],
		['PowerPointFixture.pptm', 'VBAProject'],
	])('%s is %s, as the host reported it', (file, name) => {
		expect(VbaProject.parse(openMacroContainer(fs.readFileSync(path.join(binaries, file))).vbaCfb()).name).toBe(name);
	});

	it('an Access database names its project after itself', () => {
		expect(VbaProject.parse(accessVbaCfb(fs.readFileSync(path.join(binaries, 'AccessFixture.accdb')))).name).toBe('AccessFixture');
	});
});
