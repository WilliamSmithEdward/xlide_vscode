// Issue #1415, measured in Excel, Word, PowerPoint and Access 16.0 (build
// 20430) through the VBE oracle, compile-only: expression shapes judged the
// same in every statement position. Parentheses around an aggregate, Is and
// TypeOf operands, Type fields under every operator, and the operand
// positions of Case, Print #, the file statements and single-line If arms.
// The cases themselves are in syntax_corpus/oracle/vbe_oracle_cases.json.

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
const byId = new Map(corpus.cases.map((oracleCase) => [oracleCase.id, oracleCase]));
const HOSTS = ['', 'word_', 'powerpoint_', 'access_'] as const;

/** Each error's code, analyzed as the extension analyzes the case's host; the run-time 91 of an unset Object is no syntax finding. */
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
		.filter((diagnostic) => diagnostic.severity === 'error' && diagnostic.code !== 'object-variable-not-set')
		.map((diagnostic) => diagnostic.code));
}

function ids(names: readonly string[]): Array<[string]> {
	return names.flatMap((name) => HOSTS.map((host): [string] => [`issue1415_${host}${name}_compile`]));
}

function expectOnly(id: string, expected: OracleCase['expected'], codes: string[]): void {
	const oracleCase = byId.get(id);
	if (!oracleCase) {
		throw new Error(`no oracle case ${id}`);
	}
	expect(oracleCase.expected).toBe(expected);
	expect(errorCodes(oracleCase)).toEqual(codes);
}

const OPERATORS = ['plus', 'minus', 'times', 'divide', 'intdiv', 'mod', 'concat', 'eq', 'ne', 'lt', 'gt', 'le', 'ge', 'and', 'or', 'xor', 'eqv', 'imp', 'like'];

describe('parentheses keep an array an array (issue #1415)', () => {
	it.each(ids([...OPERATORS.map((op) => `g1_paren_array_${op}`), ...['plus', 'concat', 'eq', 'and'].map((op) => `g1_double_paren_array_${op}`)]))(
		'%s is Type mismatch',
		(id) => expectOnly(id, 'rejected', ['non-scalar-binary-operand']),
	);

	it.each(ids(['g1_paren_array_pow', 'g1_paren_element_plus_ok', 'g1_paren_variant_plus_ok', 'l_element_like_ok']))('%s compiles, and nothing reports it', (id) =>
		expectOnly(id, 'accepted', []));

	it.each(ids(['l_bare_array_like', 'l_like_pattern_array']))('%s refuses the array under Like', (id) =>
		expectOnly(id, 'rejected', ['non-scalar-binary-operand']));
});

describe('Is and TypeOf operands (issue #1415)', () => {
	it.each(ids(['g2_is_paren_left', 'g2_is_paren_right', 'g2_is_double_paren', 'g2_is_element_left', 'g2_is_element_right']))(
		'%s is Type mismatch',
		(id) => expectOnly(id, 'rejected', ['is-operator-non-object']),
	);

	it.each(ids(['g3_typeof_paren', 'g3_typeof_is_long', 'g3_typeof_is_string', 'g3_typeof_is_variant', 'g3_typeof_call_is_long']))('%s is a Syntax error', (id) =>
		expectOnly(id, 'rejected', ['invalid-expression-syntax']));

	it.each(ids(['g2_is_paren_object_ok', 'g2_is_collection_element_ok', 'g3_typeof_ok', 'g3_typeof_call_ok']))('%s compiles, and nothing reports it', (id) =>
		expectOnly(id, 'accepted', []));
});

describe('a field that holds a Type, under every operator (issue #1415)', () => {
	it.each(ids(['and', 'or', 'xor', 'eqv', 'imp', 'plus', 'like'].map((op) => `g4_udt_field_${op}`)))('%s is Type mismatch', (id) =>
		expectOnly(id, 'rejected', ['non-scalar-binary-operand']));

	it.each(ids(['g4_udt_scalar_field_and_ok']))('%s compiles, and nothing reports it', (id) => expectOnly(id, 'accepted', []));
});

describe('the operand positions of statements (issue #1415)', () => {
	const ARRAYS = [
		'g5_case_array', 'g5_case_is_array', 'g5_case_is_array_plus', 'g5_print_array_plus', 'g5_print_tab_array', 'g5_print_spc_array',
		'g5_print_number_array', 'g5_write_number_array', 'g5_seek_array', 'g5_put_record_array', 'g5_get_record_array',
		'g5_line_input_number_array', 'g5_open_number_array', 'g5_close_number_array', 'g5_lock_record_array', 'g5_unlock_record_array',
		'g5_lock_number_array', 'g5_unlock_number_array', 'g5_lock_range_end_array',
	];
	it.each(ids(ARRAYS))('%s is Type mismatch', (id) => expectOnly(id, 'rejected', ['non-scalar-binary-operand']));

	it.each(ids(['g5_print_number_is', 'g5_write_number_is']))('%s is Type mismatch', (id) => expectOnly(id, 'rejected', ['is-operator-non-object']));

	it.each(ids(['g5_case_element_ok', 'g5_print_element_ok', 'g5_seek_scalar_ok', 'g5_lock_range_ok', 'g5_close_scalar_ok']))('%s compiles, and nothing reports it', (id) =>
		expectOnly(id, 'accepted', []));
});

describe('an assignment that runs on past a comma (issue #1415)', () => {
	it.each(ids(['g6_with_bang_comma', 'g6_with_dot_comma', 'h_bare_comma', 'h_member_comma', 'h_let_comma', 'h_set_comma', 'h_inline_else_comma']))(
		'%s is a Syntax error',
		(id) => expectOnly(id, 'rejected', ['invalid-expression-syntax']),
	);

	it.each(ids(['g6_inline_else_ok', 'g6_with_bang_ok', 'h_call_commas_ok', 'h_index_commas_ok', 'h_unary_minus_ok']))('%s compiles, and nothing reports it', (id) =>
		expectOnly(id, 'accepted', []));

	// A value ending at an operator is a Syntax error the VBE reports, and the
	// line someone is typing: it stays quiet (controlFlow.test.ts).
	it.each(ids(['g6_inline_else_dangling_plus', 'h_trailing_plus', 'h_inline_trailing_plus']))('%s stays quiet while it is typed', (id) => {
		expect(byId.get(id)?.expected).toBe('rejected');
		expect(errorCodes(byId.get(id)!)).not.toContain('invalid-expression-syntax');
	});
});
