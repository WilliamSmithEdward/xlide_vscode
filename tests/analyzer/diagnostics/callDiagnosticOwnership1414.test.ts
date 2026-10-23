// Issue #1414, measured in Excel, Word, PowerPoint and Access 16.0 (build
// 20430) through the VBE oracle: which error a call gets. A Function call on
// the left of an assignment is the assignment's error, a VBA.-qualified
// built-in takes the arguments its bare name does, and the plain-list
// statements check theirs. Every case is compile-only, and each statement
// that acts outside VBA sits in an If False Then block. The cases themselves
// are in syntax_corpus/oracle/vbe_oracle_cases.json.

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

interface OracleCase {
	id: string;
	expected: 'accepted' | 'rejected';
	host?: 'excel' | 'word' | 'powerpoint' | 'access';
	modules: { name: string; type?: string; source: string }[];
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

/** Each error's code, analyzed as the extension analyzes the case's host. */
function errorCodes(oracleCase: OracleCase): string[] {
	const modules: VbaProjectModuleInput[] = oracleCase.modules.map((mod) => ({ moduleName: mod.name, type: mod.type, source: mod.source }));
	const project = buildVbaProjectIndex(modules, undefined, { conditionalCompilation: { projectConstants: {} } });
	const procedures = projectProcedureSignatures(project);
	return modules.flatMap((mod) => analyzeVbaModuleSource({
		source: mod.source,
		moduleName: mod.moduleName,
		moduleKind: effectiveModuleKind(mod),
		host: oracleCase.host ?? 'excel',
		...projectAnalysisOptionsForModule(project, mod.moduleName, procedures),
	}).diagnostics
		.filter((diagnostic) => diagnostic.severity === 'error')
		.map((diagnostic) => diagnostic.code));
}

function eachHost(names: readonly string[]): Array<[string]> {
	return names.flatMap((name) => HOSTS.map((host): [string] => [`issue1414_${host}${name}_compile`]));
}

function expectOnly(id: string, expected: OracleCase['expected'], codes: string[]): void {
	const oracleCase = caseFor(id);
	expect(oracleCase.expected).toBe(expected);
	expect(errorCodes(oracleCase)).toEqual(codes);
}

describe('a Function call on the left of an assignment (issue #1414)', () => {
	it.each(eachHost(['t_paren', 't_let', 't_qualified', 't_let_qualified', 't_with_argument', 't_optional', 't_bare']))(
		'%s is the assignment error, and only that',
		(id) => expectOnly(id, 'rejected', ['assignment-to-procedure-name']),
	);

	it.each(eachHost(['t_variant_missing_argument']))('%s returns a Variant, so its argument is what is wrong', (id) =>
		expectOnly(id, 'rejected', ['argument-count']));

	it.each(eachHost(['t_variant_with_argument']))('%s compiles, and nothing reports it', (id) => expectOnly(id, 'accepted', []));
});

describe('a VBA.-qualified built-in takes the arguments its bare name does (issue #1414)', () => {
	const CALLS = ['clng_none', 'cint_none', 'left_dollar_none', 'len_none', 'clng_extra', 'instr_five'];
	it.each(eachHost(CALLS.flatMap((call) => [`q_bare_${call}`, `q_vba_${call}`])))('%s has the wrong arguments', (id) =>
		expectOnly(id, 'rejected', ['argument-count']));

	it.each(eachHost(['q_vba_clng_ok', 'q_vba_left_dollar_ok', 'q_vba_len_ok', 'q_vba_instr_four_ok']))(
		'%s compiles, and nothing reports it',
		(id) => expectOnly(id, 'accepted', []),
	);
});

describe('the plain-list statements check their arguments (issue #1414)', () => {
	const STATEMENTS = ['savesetting', 'deletesetting', 'sendkeys', 'chdir', 'chdrive', 'mkdir', 'rmdir', 'kill', 'setattr', 'filecopy', 'appactivate'];
	const REFUSED = [
		...STATEMENTS.flatMap((name) => [`s2_${name}_missing`, `s2_${name}_extra`]),
		's2_savesetting_duplicate_named', 's_appactivate_omitted_title', 's_appactivate_duplicate_named', 's_filecopy_omitted_source',
	];
	it.each(eachHost(REFUSED))('%s has the wrong arguments', (id) => expectOnly(id, 'rejected', ['argument-count']));

	it.each(eachHost([...STATEMENTS.map((name) => `s2_${name}_valid`), 's_appactivate_ok', 's_appactivate_named_ok', 's_filecopy_ok']))(
		'%s compiles, and nothing reports it',
		(id) => expectOnly(id, 'accepted', []),
	);
});
