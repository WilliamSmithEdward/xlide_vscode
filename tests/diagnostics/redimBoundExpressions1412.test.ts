// Issue #1412, measured in Excel, Word, PowerPoint and Access 16.0 (build
// 20430) through the VBE oracle, compile-only: a ReDim bound is an
// expression, judged as one. `ReDim d(n Is Nothing)` with n a Long and
// `ReDim d(a)` with a an array are Type mismatch. The cases are in
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
	return names.flatMap((name) => HOSTS.map((host): [string] => [`issue1412_${host}${name}_compile`]));
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

const main = (...body: string[]): string => ['Option Explicit', 'Sub Main()', '    Dim a(1) As Long, d() As Long, n As Long', ...body, 'End Sub', ''].join('\r\n');

describe('a ReDim bound is an expression (#1412)', () => {
	it.each(ids(['rb_is_nothing', 'rb_cdbl_is', 'rb_typeof_n']))('%s is Type mismatch, through the Is operand rule', (id) => {
		expectOnly(id, 'rejected', ['is-operator-non-object']);
	});

	it.each(ids(['rb_array_upper', 'rb_array_lower', 'rb_array_paren', 'rb_array_preserve', 'rb_array_second_dim']))('%s is an array as a bound', (id) => {
		expectOnly(id, 'rejected', ['redim-bound-array']);
	});

	it.each(ids(['rb_array_plus']))('%s is an array under an operator', (id) => {
		expectOnly(id, 'rejected', ['non-scalar-binary-operand']);
	});

	it.each(ids(['rb_element_ok', 'rb_ubound_ok', 'rb_variant_ok', 'rb_object_is_ok', 'rb_two_dims_ok']))('%s compiles', (id) => {
		expectOnly(id, 'accepted', []);
	});

	it('reads each variable of a ReDim list, and leaves its As type alone', () => {
		const source = main('    Dim e() As Long', '    ReDim d(1) As Long, e(a)');
		expect(errorCodes([{ moduleName: 'Module1', source }])).toEqual(['redim-bound-array']);
	});

	it('reports a ReDim in an If arm through the operand rules only', () => {
		const source = main('    If n = 0 Then ReDim d(n Is Nothing)');
		expect(errorCodes([{ moduleName: 'Module1', source }])).toEqual(['is-operator-non-object']);
	});
});
