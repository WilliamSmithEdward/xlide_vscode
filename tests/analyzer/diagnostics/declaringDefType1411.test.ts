// Issue #1411, measured in Excel, Word, PowerPoint and Access 16.0 (build
// 20430) through the VBE oracle: an untyped procedure of another module takes
// its types from that module's DefType lines, whatever the caller's say; and
// Is and TypeOf refuse a scalar operand, a call result included. The cases
// themselves are in syntax_corpus/oracle/vbe_oracle_cases.json, one per host.

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
	return HOSTS.map((host) => [`issue1411_${host}${name}_${phase}`]);
}

describe("an untyped parameter takes its declaring module's DefType (issue #1411)", () => {
	it.each(eachHost('f_byref_string_variable', 'compile'))('%s passes a String ByRef to a Long', (id) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('rejected');
		expect(errorsFor(oracleCase)).toEqual(['byref-argument-type-mismatch: s']);
	});

	it.each(eachHost('f_text_string_literal', 'runtime'))('%s passes text that is no number to a Long', (id) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('rejected');
		expect(errorsFor(oracleCase)).toEqual(['argument-type-mismatch: "abc"']);
	});

	it.each([...eachHost('f_numeric_string_literal', 'runtime'), ...eachHost('f_caller_deftype_only', 'compile')])('%s runs, and nothing reports it', (id) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('accepted');
		expect(errorsFor(oracleCase)).toEqual([]);
	});
});

describe('Is and TypeOf refuse a scalar operand (issue #1411)', () => {
	const REFUSED: Array<[string, string]> = [
		['g_is_long_function', 'Bar()'],
		['g_is_long_function_with_argument', 'Bar(1)'],
		['g_is_string_function', 'Bar()'],
		['g_is_long_array_function', 'Bar()'],
		['g_is_deflng_bare', 'Foo()'],
		['g_is_deflng_qualified', 'Other.Foo()'],
		['g_is_deflng_caller_defobj', 'Foo()'],
		['g_typeof_long_identifier', 'n'],
		['g_typeof_string_identifier', 's'],
		['g_typeof_long_function', 'Bar()'],
		['g_typeof_deflng_qualified', 'Other.Foo()'],
	];
	it.each(REFUSED.flatMap(([name, operand]) => eachHost(name, 'compile').map(([id]) => [id, operand])))('%s refuses %s', (id, operand) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('rejected');
		expect(errorsFor(oracleCase)).toEqual([`is-operator-non-object: ${operand}`]);
	});

	const ACCEPTED = [
		'g_is_object_function', 'g_is_variant_function', 'g_is_defobj_qualified',
		'g_typeof_variant_identifier', 'g_typeof_object_identifier', 'g_typeof_defobj_bare',
	];
	it.each(ACCEPTED.flatMap((name) => eachHost(name, 'compile')))('%s compiles, and nothing reports it', (id) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('accepted');
		expect(errorsFor(oracleCase)).toEqual([]);
	});

	// The local Other is the receiver, and is Nothing when it runs: the 91 is
	// right, and Is has nothing to say about a late-bound call.
	it.each(eachHost('g_is_deflng_shadowed_qualifier', 'compile'))('%s leaves a late-bound call alone', (id) => {
		const oracleCase = caseFor(id);
		expect(oracleCase.expected).toBe('accepted');
		expect(errorsFor(oracleCase).filter((finding) => finding.startsWith('is-operator-non-object:'))).toEqual([]);
	});
});
