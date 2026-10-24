// The Outlook object library as a project reference: an Excel, Word,
// PowerPoint or Access project that references it names Outlook's types,
// constants and global members early bound. Measured in all four hosts
// (16.0, build 20430) through the VBE oracle, with the reference added by
// GUID; the cases are in syntax_corpus/oracle/vbe_oracle_cases.json.

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { resolveMemberCompletions } from '../../src/analyzer/completion/memberAccess';
import { resolveTypeCompletions } from '../../src/analyzer/completion/typeCompletion';
import { hostObjectModelForTokens } from '../../src/analyzer/host/hostRegistry';
import { resolveHover } from '../../src/analyzer/hover/resolveHover';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';
import {
	buildVbaProjectIndex,
	effectiveModuleKind,
	projectAnalysisOptionsForModule,
	projectProcedureSignatures,
} from '../../src/vbaProjectAnalysis';
import { oracleModuleInput, oracleReferencedHosts, type OracleFixtureModule } from '../helpers/oracleModuleInput';

interface OracleCase {
	id: string;
	expected: 'accepted' | 'rejected';
	host?: 'excel' | 'word' | 'powerpoint' | 'access';
	references?: { guid: string }[];
	modules: OracleFixtureModule[];
}

const corpus: { cases: OracleCase[] } = JSON.parse(
	readFileSync(join(process.cwd(), 'syntax_corpus', 'oracle', 'vbe_oracle_cases.json'), 'utf8'),
);
const HOSTS = ['', 'word_', 'powerpoint_', 'access_'] as const;

function caseFor(id: string): OracleCase {
	const found = corpus.cases.find((oracleCase) => oracleCase.id === id);
	if (!found) {
		throw new Error(`no oracle case ${id}`);
	}
	return found;
}

/**
 * Each error code, analyzed as the extension analyzes the case's host and
 * references. `referencedLibraries` is the project's reference list where
 * the extension read it; undefined leaves the rules that need it silent.
 */
function errorCodes(oracleCase: OracleCase, referencedLibraries?: readonly string[]): string[] {
	const modules = oracleCase.modules.map(oracleModuleInput);
	const project = buildVbaProjectIndex(modules, undefined, { conditionalCompilation: { projectConstants: {} } });
	const procedures = projectProcedureSignatures(project);
	return modules.flatMap((mod) => analyzeVbaModuleSource({
		source: mod.source,
		moduleName: mod.moduleName,
		moduleKind: effectiveModuleKind(mod),
		host: oracleCase.host ?? 'excel',
		...oracleReferencedHosts(oracleCase),
		...(referencedLibraries ? { referencedLibraries } : {}),
		...projectAnalysisOptionsForModule(project, mod.moduleName, procedures),
	}).diagnostics.filter((diagnostic) => diagnostic.severity === 'error').map((diagnostic) => diagnostic.code ?? ''));
}

function eachHost(name: string, hosts: readonly string[] = HOSTS): Array<[string]> {
	return hosts.map((host) => [`outlook_${host}${name}_compile`]);
}

describe('a project that references Outlook (measured)', () => {
	it.each([
		'ref_dim_app', 'ref_dim_mailitem', 'ref_dim_bare_mailitem', 'ref_new_app', 'ref_constant', 'ref_qualified_constant',
		'ref_enum_constant', 'ref_member_ok', 'ref_member_missing', 'ref_item_member_ok', 'ref_item_member_missing',
		'ref_bare_session', 'ref_bare_createitem', 'ref_bare_getnamespace', 'ref_module_level', 'ref_withevents',
	].flatMap((name) => eachHost(name)))('%s compiles, and nothing reports it', (id) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('accepted');
		expect(oracleCase.references).toHaveLength(1);
		expect(errorCodes(oracleCase)).toEqual([]);
	});

	it.each([
		['noref_dim_app', 'missing-library-reference'],
		['noref_constant', 'undeclared-variable'],
		['ref_bare_unknown', 'undeclared-variable'],
		['ref_wrong_arguments', 'argument-count'],
	].flatMap(([name, code]) => eachHost(name).map(([id]): [string, string] => [id, code])))('%s is refused, as %s', (id, code) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('rejected');
		expect(errorCodes(oracleCase)).toEqual([code]);
	});

	// A bare type name is judged against the libraries the project is known
	// to reference, so it waits for the reference list.
	it.each(eachHost('noref_dim_bare_mailitem'))('%s is refused once the references are known to leave Outlook out', (id) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('rejected');
		expect(errorCodes(oracleCase)).toEqual([]);
		expect(errorCodes(oracleCase, ['stdole', 'Office'])).toEqual(['invalid-as-type-name']);
	});

	it.each(eachHost('ref_dim_bare_mailitem'))('%s resolves a bare MailItem once the list holds Outlook', (id) => {
		expect(errorCodes(caseFor(id), ['stdole', 'Office', 'Outlook'])).toEqual([]);
	});
});

describe("a referenced library's global object binds bare after the host's (measured)", () => {
	const others = HOSTS.filter((host) => host !== '');

	it.each(eachHost('xlref_bare_activeworkbook', others))("%s: Excel's ActiveWorkbook, with Excel referenced", (id) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('accepted');
		expect(errorCodes(oracleCase)).toEqual([]);
	});

	it.each(eachHost('noxlref_bare_activeworkbook', others))('%s: and Variable not defined without it', (id) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('rejected');
		expect(errorCodes(oracleCase)).toEqual(['undeclared-variable']);
	});

	it("keeps the host's own global ahead of a referenced library's", () => {
		const source = 'Sub Main()\r\n    Dim x As Object\r\n    Set x = Application\r\nEnd Sub\r\n';
		const hover = resolveHover(source, source.indexOf('Application') + 2, { model: hostObjectModelForTokens(['excel', 'outlook']) });
		expect(hover?.signature).toContain('Application');
		expect(JSON.stringify(hover)).not.toContain('Outlook');
	});
});

describe('the Outlook model in the editor', () => {
	const model = hostObjectModelForTokens(['excel', 'outlook']);

	it("offers Outlook's members on an Outlook object", () => {
		const source = 'Sub Main()\r\n    Dim o As Outlook.Application\r\n    o.\r\nEnd Sub\r\n';
		const names = resolveMemberCompletions(source, source.indexOf('o.\r\n') + 2, { model }).map((item) => item.name);
		expect(names).toEqual(expect.arrayContaining(['GetNamespace', 'CreateItem', 'Session']));
	});

	it('offers Outlook types after As, only when the project references it', () => {
		const source = 'Dim m As \r\n';
		expect(resolveTypeCompletions(source, 9, { model }).map((item) => item.name)).toContain('MailItem');
		expect(resolveTypeCompletions(source, 9, { model: hostObjectModelForTokens(['excel']) }).map((item) => item.name)).not.toContain('MailItem');
	});

	it('describes a bare Outlook global', () => {
		const source = 'Sub Main()\r\n    Dim x As Object\r\n    Set x = Session\r\nEnd Sub\r\n';
		expect(JSON.stringify(resolveHover(source, source.indexOf('Session') + 2, { model }))).toContain('NameSpace');
	});
});
