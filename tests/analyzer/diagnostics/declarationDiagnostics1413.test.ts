// Issue #1413, measured in Excel, Word, PowerPoint and Access 16.0 (build
// 20430) through the VBE oracle, compile-only: Const Nothing, a Declare's
// return type, Tab as a type, a fixed-length String result and New with an
// argument list. The cases themselves are in
// syntax_corpus/oracle/vbe_oracle_cases.json.

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

/** Each error as `code: spanned text @ offset`, analyzed as the extension analyzes the case's host. */
function errorsFor(oracleCase: OracleCase): Array<{ code: string; text: string; start: number; source: string }> {
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
		.map((diagnostic) => ({ code: diagnostic.code, text: mod.source.slice(diagnostic.span.start, diagnostic.span.end), start: diagnostic.span.start, source: mod.source })));
}

function ids(names: readonly string[], hosts: readonly string[] = HOSTS): Array<[string]> {
	return names.flatMap((name) => hosts.map((host): [string] => [`issue1413_${host}${name}_compile`]));
}

function expectCodes(id: string, expected: OracleCase['expected'], codes: string[]): void {
	const oracleCase = caseFor(id);
	expect(oracleCase.expected).toBe(expected);
	expect(errorsFor(oracleCase).map((error) => error.code)).toEqual(codes);
}

describe('a Const given Nothing (issue #1413)', () => {
	it.each(ids(['c_single', 'c_single_suffix', 'c_double_parenthesized', 'c_long_double_parenthesized', 'c_longlong', 'c_longptr', 'c_string', 'c_untyped', 'c_variant', 'c_module_level']))(
		'%s is refused at Nothing',
		(id) => {
			const oracleCase = caseFor(id);
			expect(oracleCase.expected).toBe('rejected');
			const errors = errorsFor(oracleCase);
			expect(errors.map((error) => `${error.code}: ${error.text}`)).toEqual(['const-value-not-constant: Nothing']);
		},
	);

	it('gives each type the message the VBE gives', () => {
		const message = (name: string) => analyzeVbaModuleSource({
			source: caseFor(`issue1413_${name}_compile`).modules[0].source,
			moduleName: 'Module1',
			moduleKind: 'standard',
			host: 'excel',
		}).diagnostics.find((diagnostic) => diagnostic.code === 'const-value-not-constant')?.message;
		expect(message('c_single')).toContain('Invalid use of object');
		expect(message('c_untyped')).toContain('Invalid data type for constant');
		expect(message('c_variant')).toContain('Object variable or With block variable not set');
	});

	it.each(ids(['c_single_ok']))('%s compiles, and nothing reports it', (id) => expectCodes(id, 'accepted', []));
});

describe("a Declare's return type without its reference (issue #1413)", () => {
	const OTHER_HOSTS = ['word_', 'powerpoint_', 'access_'];
	it.each(ids(['x_declare_spaced_type', 'x_declare_unspaced_type', 'x_dim_spaced_type'], OTHER_HOSTS))('%s names Excel without the reference', (id) =>
		expectCodes(id, 'rejected', ['missing-library-reference']));

	it.each(ids(['x_declare_type_text_in_lib'], OTHER_HOSTS))('%s is reported at the type, not inside the Lib string', (id) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('rejected');
		const [error, ...rest] = errorsFor(oracleCase);
		expect(rest).toEqual([]);
		expect(error.code).toBe('missing-library-reference');
		expect(error.start).toBeGreaterThan(error.source.indexOf(' As '));
	});

	it.each(ids(['x_lib_text_long_ok'], OTHER_HOSTS))('%s compiles, and nothing reports it', (id) => expectCodes(id, 'accepted', []));
});

describe('Tab and Spc are no type names (issue #1413)', () => {
	it.each(ids(['t_dim_tab', 't_dim_spc', 't_param_tab']))('%s is refused', (id) => expectCodes(id, 'rejected', ['invalid-as-type-name']));

	it.each([...ids(['t_name_tab_ok']), ['issue1413_t_dim_excel_tab_ok_compile'] as [string]])('%s compiles, and nothing reports it', (id) =>
		expectCodes(id, 'accepted', []));
});

describe('no procedure returns a fixed-length String (issue #1413)', () => {
	it.each(ids(['f_no_parens_fixed_string', 'f_parens_fixed_string']))('%s is refused at the *', (id) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('rejected');
		expect(errorsFor(oracleCase).map((error) => `${error.code}: ${error.text}`)).toEqual(['unexpected-declaration-token: *']);
	});

	it.each(ids(['f_no_parens_string_ok']))('%s compiles, and nothing reports it', (id) => expectCodes(id, 'accepted', []));
});

describe('New takes no argument list (issue #1413)', () => {
	it.each([...ids(['n_class_parens', 'n_collection_parens']), ['issue1413_n_excel_application_parens_compile'] as [string]])('%s is refused', (id) =>
		expectCodes(id, 'rejected', ['invalid-expression-syntax']));

	it.each(ids(['n_dim_as_new_parens']))('%s stays the declaration error', (id) => expectCodes(id, 'rejected', ['unexpected-declaration-token']));

	it.each(ids(['n_class_ok']))('%s compiles, and nothing reports it', (id) => expectCodes(id, 'accepted', []));
});
