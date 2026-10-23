// Issue #1412, measured in Excel, Word, PowerPoint and Access 16.0 (build
// 20430) through the VBE oracle: a ReDim's element type compared by which
// type it names, not by its spelling, and bounds a Long cannot hold. The
// cases themselves are in syntax_corpus/oracle/vbe_oracle_cases.json.

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

// A bound is made whole half to even, then must fit a Long: 2147483647.5 is
// 2147483648 and overflows, 2147483647.4 and -2147483648.5 do not. A Double
// Const, and an expression of Doubles, is folded with VBA's typing.
describe('array bounds folded past whole numbers (issue #1412)', () => {
	const DIM_OVERFLOWS = [
		'l_dim_double_overflow', 'd_half_up', 'd_low_past', 'd_low_double', 'd_const_double', 'd_local_const_double',
		'd_const_untyped', 'd_const_plus', 'd_double_plus', 'd_divide', 'd_single', 'd_currency', 'd_second_dim',
		'd_module_level', 'd_type_field', 'd_long_sum',
	];
	it.each(DIM_OVERFLOWS.flatMap((name) => eachHost(name, 'compile')))('%s is a Dim bound a Long cannot hold', (id) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('rejected');
		expect(errorsFor(oracleCase).map((finding) => finding.split(':')[0])).toEqual(['array-bound-overflow']);
	});

	const REDIM_OVERFLOWS = [
		'l_redim_const_overflow', 'r_half_up', 'r_low_past', 'r_local_const', 'r_const_plus', 'r_double_plus',
		'r_divide', 'r_single', 'r_currency',
	];
	it.each(REDIM_OVERFLOWS.flatMap((name) => eachHost(name, 'runtime')))('%s raises Overflow at the ReDim', (id) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('rejected');
		expect(errorsFor(oracleCase).map((finding) => finding.split(':')[0])).toEqual(['arithmetic-overflow']);
	});

	it.each([
		...['d_under_half_ok', 'd_low_half_ok', 'd_const_ok', 'd_divide_ok'].flatMap((name) => eachHost(name, 'compile')),
		...['r_under_half_ok', 'r_low_half_ok', 'r_const_ok'].flatMap((name) => eachHost(name, 'runtime')),
	])('%s fits a Long, and nothing reports it', (id) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('accepted');
		expect(errorsFor(oracleCase)).toEqual([]);
	});

	const moduleOf = (...lines: string[]): OracleCase => ({
		id: 'unit',
		expected: 'rejected',
		modules: [{ name: 'Module1', source: ['Option Explicit', ...lines, ''].join('\r\n') }],
	});

	it('names the bound, its dimension and the compile error, and underlines the dimension', () => {
		const oracleCase = moduleOf('Const N As Double = 2147483647#', 'Sub Main()', '    Dim a(1, 0 To N + 1) As Long', 'End Sub');
		expect(errorsFor(oracleCase)).toEqual(['array-bound-overflow: 0 To N + 1']);
		const source = oracleCase.modules[0].source;
		const project = buildVbaProjectIndex([{ moduleName: 'Module1', source }], undefined, { conditionalCompilation: { projectConstants: {} } });
		const message = analyzeVbaModuleSource({
			source,
			moduleName: 'Module1',
			moduleKind: 'standard',
			...projectAnalysisOptionsForModule(project, 'Module1', projectProcedureSignatures(project)),
		}).diagnostics.find((diagnostic) => diagnostic.code === 'array-bound-overflow')?.message;
		expect(message).toBe("Array 'a' bound 2147483648 for dimension 2 is outside the range of a Long; this is a VBE compile error: Overflow.");
	});

	it('reports a bound past the range once, not also as reversed bounds', () => {
		expect(errorsFor(moduleOf('Sub Main()', '    Dim a(2147483648# To 1) As Long', 'End Sub')))
			.toEqual(['array-bound-overflow: 2147483648# To 1']);
	});

	it('leaves a suffixed literal out of range to suffixed-literal-overflow', () => {
		expect(errorsFor(moduleOf('Sub Main()', '    Dim a(2147483648&) As Long', 'End Sub')).map((finding) => finding.split(':')[0]))
			.toEqual(['suffixed-literal-overflow']);
	});

	it('reads a Const of another module, and leaves a local that hides a Const to the constant rule', () => {
		const shared = { name: 'Limits', source: 'Option Explicit\r\nPublic Const BIG As Double = 3E+09\r\n' };
		const user: OracleCase = { id: 'unit', expected: 'rejected', modules: [shared, { name: 'Module1', source: 'Option Explicit\r\nSub Main()\r\n    Dim a(BIG) As Long\r\nEnd Sub\r\n' }] };
		expect(errorsFor(user)).toEqual(['array-bound-overflow: BIG']);
		const hidden = moduleOf('Const BIG As Double = 3E+09', 'Sub Main()', '    Dim BIG As Long', '    Dim a(BIG) As Long', 'End Sub');
		expect(errorsFor(hidden).map((finding) => finding.split(':')[0])).not.toContain('array-bound-overflow');
	});
});
