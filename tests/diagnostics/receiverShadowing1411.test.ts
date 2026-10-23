// Issue #1411, measured in Excel, Word, PowerPoint and Access 16.0 (build
// 20430) through the VBE oracle: a standard module's name used as a
// receiver while a variable takes that name, and a module variable reached
// through its module after it was set to Nothing. The cases themselves are
// in syntax_corpus/oracle/vbe_oracle_cases.json, one per host.

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';
import {
	buildVbaProjectIndex,
	effectiveModuleKind,
	projectAnalysisOptionsForModule,
	projectProcedureSignatures,
	type VbaProjectModuleInput,
} from '../../src/vbaProjectAnalysis';

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

/** Each finding as `code: spanned text`, analyzed as the extension analyzes the case's host. */
function findingsFor(oracleCase: OracleCase): string[] {
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
		.map((diagnostic) => `${diagnostic.code}: ${mod.source.slice(diagnostic.span.start, diagnostic.span.end)}`));
}

function eachHost(name: string, phase: 'compile' | 'runtime'): Array<[string]> {
	return HOSTS.map((host) => [`issue1411_${host}${name}_${phase}`]);
}

describe('a module name a variable takes is that variable (issue #1411)', () => {
	const SHADOWED = [
		...['strings', 'math', 'other'].flatMap((receiver) => ['local', 'parameter', 'module', 'variant_local', 'current_module']
			.filter((by) => by !== 'current_module' || receiver !== 'other')
			.map((by) => `b_${receiver}_${by}`)),
		'b_other_collection_local',
		...['statement', 'call_statement', 'assignment', 'nested_argument', 'set_assignment'].map((shape) => `d_${shape}_shadowed`),
	];
	it.each(SHADOWED.flatMap((name) => eachHost(name, 'compile')))('%s compiles, and nothing reports it', (id) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('accepted');
		expect(findingsFor(oracleCase)).toEqual([]);
	});

	const UNSHADOWED = [
		'b_strings_unshadowed', 'b_math_unshadowed', 'b_other_unshadowed',
		...['statement', 'call_statement', 'assignment', 'nested_argument', 'set_assignment'].map((shape) => `d_${shape}_unshadowed`),
		// A Public variable of a third module does not take the name.
		'd_public_variable_named_like_module',
	];
	it.each(UNSHADOWED.flatMap((name) => eachHost(name, 'compile')))('%s is still a call to the module, with an argument missing', (id) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('rejected');
		expect(findingsFor(oracleCase).filter((finding) => finding.startsWith('argument-count:'))).toHaveLength(1);
	});

	it.each(eachHost('b_other_long_local', 'compile'))('%s is a member of a Long, not a call to the module', (id) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('rejected');
		const findings = findingsFor(oracleCase);
		expect(findings.some((finding) => finding.startsWith('scalar-member-access:'))).toBe(true);
		expect(findings.some((finding) => finding.startsWith('argument-count:'))).toBe(false);
	});
});

describe('a module variable reset to Nothing through its module (issue #1411)', () => {
	const READS = ['print', 'call', 'label', 'if', 'assign'].map((form) => `a_nothing_read_${form}`);
	it.each([...READS, 'a_reset_unqualified_read_qualified'].flatMap((name) => eachHost(name, 'runtime')))('%s raises 91 at the qualified read', (id) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('rejected');
		expect(findingsFor(oracleCase)).toEqual(['object-variable-not-set: Module2.g']);
	});

	it.each(eachHost('a_nothing_read_with', 'runtime'))('%s raises 91 at the With', (id) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('rejected');
		expect(findingsFor(oracleCase)).toEqual(['object-variable-not-set: Module2.g']);
	});

	const RUNS = [
		'a_reset_then_new', 'a_reset_then_call_init', 'a_shadowed_own_module_receiver',
		'e_block_guard', 'e_line_guard', 'e_exit_between', 'e_as_new',
	];
	it.each(RUNS.flatMap((name) => eachHost(name, 'runtime')))('%s runs, and nothing reports it', (id) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('accepted');
		expect(findingsFor(oracleCase)).toEqual([]);
	});
});
