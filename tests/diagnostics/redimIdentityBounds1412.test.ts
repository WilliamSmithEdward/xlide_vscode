// Issue #1412, measured in Excel, Word, PowerPoint and Access 16.0 (build
// 20430) through the VBE oracle: a ReDim's element type compared by which
// type it names, not by its spelling, and bounds a Long cannot hold. The
// cases themselves are in syntax_corpus/oracle/vbe_oracle_cases.json.

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

/** Each error as `code: spanned text`, analyzed as the extension analyzes the case's host. */
function errorsFor(oracleCase: OracleCase): string[] {
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
	return HOSTS.map((host) => [`issue1412_${host}${name}_${phase}`]);
}

describe('a ReDim compares the type its element type names (issue #1412)', () => {
	const CHANGES: Array<[string, string]> = [
		...['r_project_collection_to_vba', 'r_vba_collection_to_project', 'r_long_to_integer']
			.flatMap((name) => eachHost(name, 'compile').map(([id]): [string, string] => [id, ''])),
		['issue1412_r_project_range_to_excel_compile', ''],
		['issue1412_word_r_project_range_to_word_compile', ''],
	];
	it.each(CHANGES)('%s changes the element type', (id) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('rejected');
		expect(errorsFor(oracleCase).filter((finding) => finding.startsWith('redim-type-change:'))).toHaveLength(1);
	});

	const SAME: Array<[string]> = [
		...['r_project_collection_same', 'r_collection_to_vba_no_class'].flatMap((name) => eachHost(name, 'compile')),
		['issue1412_r_range_to_excel_no_class_compile'],
		['issue1412_word_r_range_to_word_no_class_compile'],
	];
	it.each(SAME)('%s names the same type, and nothing reports it', (id) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('accepted');
		expect(errorsFor(oracleCase)).toEqual([]);
	});
});

describe('array bounds a Long cannot hold (issue #1412)', () => {
	const LITERALS = [
		'l_dim_explicit_lower_long_overflow', 'l_dim_upper_long_overflow', 'l_redim_long_overflow',
		'l_redim_explicit_lower_long_overflow', 'l_dim_integer_suffix_overflow',
	];
	it.each(LITERALS.flatMap((name) => eachHost(name, 'compile')))('%s is the malformed literal, and only that', (id) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('rejected');
		const errors = errorsFor(oracleCase);
		expect(errors.map((finding) => finding.split(':')[0])).toEqual(['suffixed-literal-overflow']);
	});

	const OVERFLOWS = [
		'l_redim_double_overflow', 'l_redim_unsuffixed_overflow',
		'l_redim_explicit_upper_double_overflow', 'l_redim_negative_lower_double_overflow',
	];
	it.each(OVERFLOWS.flatMap((name) => eachHost(name, 'runtime')))('%s raises Overflow at the ReDim', (id) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('rejected');
		expect(errorsFor(oracleCase).map((finding) => finding.split(':')[0])).toEqual(['arithmetic-overflow']);
	});

	it.each([
		...eachHost('l_dim_long_max', 'compile'),
		...['l_redim_double_in_range', 'b_element_bound', 'b_long_bound'].flatMap((name) => eachHost(name, 'runtime')),
	])('%s runs, and nothing reports it', (id) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('accepted');
		expect(errorsFor(oracleCase)).toEqual([]);
	});
});
