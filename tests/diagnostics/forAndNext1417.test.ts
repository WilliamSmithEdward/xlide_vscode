// Issue #1417, measured in Excel, Word, PowerPoint and Access 16.0 (build
// 20430) through the VBE oracle, compile-only: a Next given a subscript, an
// array as a For start, end or Step value, and an Is in a single-line If
// nested in another's arm. The cases themselves are in
// syntax_corpus/oracle/vbe_oracle_cases.json.

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

/** Each error's code, analyzed as the extension analyzes the case's host. */
function errorCodes(modules: VbaProjectModuleInput[], host: OracleCase['host'] = 'excel'): string[] {
	const project = buildVbaProjectIndex(modules, undefined, { conditionalCompilation: { projectConstants: {} } });
	const procedures = projectProcedureSignatures(project);
	return modules.flatMap((mod) => analyzeVbaModuleSource({
		source: mod.source,
		moduleName: mod.moduleName,
		moduleKind: effectiveModuleKind(mod),
		host,
		...projectAnalysisOptionsForModule(project, mod.moduleName, procedures),
	}).diagnostics
		.filter((diagnostic) => diagnostic.severity === 'error')
		.map((diagnostic) => diagnostic.code));
}

function ids(names: readonly string[]): Array<[string]> {
	return names.flatMap((name) => HOSTS.map((host): [string] => [`issue1417_${host}${name}_compile`]));
}

function expectOnly(id: string, expected: OracleCase['expected'], codes: string[]): void {
	const oracleCase = byId.get(id);
	if (!oracleCase) {
		throw new Error(`no oracle case ${id}`);
	}
	expect(oracleCase.expected).toBe(expected);
	const modules = oracleCase.modules.map((mod) => ({ moduleName: mod.name, type: mod.type, source: mod.source }));
	expect(errorCodes(modules, oracleCase.host)).toEqual(codes);
}

const main = (...body: string[]): string => ['Option Explicit', 'Sub Main()', ...body, 'End Sub', ''].join('\r\n');

describe('a Next given a subscript (#1417)', () => {
	it.each(ids([
		'n_next_index_long', 'n_next_index_integer', 'n_next_index_byte', 'n_next_index_date',
		'n_next_index_single', 'n_next_index_double', 'n_next_index_currency', 'n_next_index_longlong',
		'n_next_index_longptr', 'n_next_index_long_suffix', 'n_next_index_integer_suffix',
	]))('%s is Expected array', (id) => {
		expectOnly(id, 'rejected', ['scalar-indexed']);
	});

	it.each(ids(['n_next_index_variant']))('%s is Invalid Next control variable reference', (id) => {
		expectOnly(id, 'rejected', ['next-variable-mismatch']);
	});

	it.each(ids(['n_next_index_boolean']))('%s is the counter type first', (id) => {
		expectOnly(id, 'rejected', ['for-counter-type']);
	});

	it.each(ids(['n_next_ok', 'n_next_suffix_ok']))('%s compiles', (id) => {
		expectOnly(id, 'accepted', []);
	});

	it('reports once for a loop inside an If block', () => {
		const source = main('    Dim i As Long', '    If True Then', '        For i = 1 To 2', '        Next i(1)', '    End If');
		expect(errorCodes([{ moduleName: 'Module1', source }])).toEqual(['scalar-indexed']);
	});
});

describe('an array as a For start, end or Step value (#1417)', () => {
	it.each(ids([
		'f_lineno_array_start', 'f_lineno_array_end', 'f_lineno_array_step', 'f_array_start',
		'f2_dynamic_array_end', 'f2_variant_array_start', 'f2_variant_counter_end', 'f2_paren_array_end',
		'f2_string_array_end',
	]))('%s is Type mismatch', (id) => {
		expectOnly(id, 'rejected', ['for-bound-array']);
	});

	it.each(ids(['f_lineno_ok', 'f2_variant_holding_ok', 'f2_element_ok', 'f2_ubound_ok']))('%s compiles', (id) => {
		expectOnly(id, 'accepted', []);
	});
});

describe('Is in a single-line If nested in an arm (#1417)', () => {
	it.each(ids(['e_else_if_is', 'e_if_is', 'e2_then_if_is']))('%s is Type mismatch', (id) => {
		expectOnly(id, 'rejected', ['is-operator-non-object']);
	});

	it.each(ids(['e_else_if_ok', 'e2_nested_else_if_ok']))('%s compiles', (id) => {
		expectOnly(id, 'accepted', []);
	});
});
