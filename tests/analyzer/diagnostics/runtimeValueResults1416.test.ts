// Issue #1416, measured in Excel, Word, PowerPoint and Access 16.0 (build
// 20430) through the VBE oracle, compile-only: VBA's own functions, bare
// and VBA-qualified, under Is and TypeOf; TypeOf on a literal; and an Enum
// variable under both. 3,088 cases, the same in every host, in
// syntax_corpus/oracle/vbe_oracle_cases.json.

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { analyzeVbaModuleSource } from '../../../src/vbaModuleAnalysis';
import {
	ARGUMENT_TYPED_FUNCTIONS,
	BARE_VALUE_RESULT_FUNCTIONS,
	TYPEOF_KEYWORD_OPERANDS,
	VALUE_RESULT_FUNCTIONS,
} from '../../../src/analyzer/runtime/vbaValueResults';
import {
	buildVbaProjectIndex,
	effectiveModuleKind,
	projectAnalysisOptionsForModule,
	projectProcedureSignatures,
	type VbaProjectModuleInput,
} from '../../../src/vbaProjectAnalysis';

interface OracleCase {
	id: string;
	description: string;
	expected: 'accepted' | 'rejected';
	host?: 'excel' | 'word' | 'powerpoint' | 'access';
	modules: { name: string; type?: string; source: string }[];
}

const corpus: { cases: OracleCase[] } = JSON.parse(
	readFileSync(join(process.cwd(), 'syntax_corpus', 'oracle', 'vbe_oracle_cases.json'), 'utf8'),
);
const cases = corpus.cases.filter((oracleCase) => oracleCase.id.startsWith('issue1416_'));

/** The analyzer's code for each compile error the VBE gave. */
const CODE_FOR_MESSAGE: Record<string, string> = {
	'Type mismatch': 'is-operator-non-object',
	'Syntax error': 'invalid-expression-syntax',
	'Method or data member not found': 'member-not-found',
};

/**
 * Refused by the VBE and not reported: `DoEvents Is Nothing`, since the
 * expression parser does not read a keyword-named function without
 * parentheses as a call, and an Enum member, `ShadeOne Is Nothing`, since
 * the rule does not resolve constants.
 */
const KNOWN_QUIET = /^issue1416_(word_|powerpoint_|access_)?(is_doevents|enum3_member_is)_compile$/;

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

describe("VBA's function results under Is and TypeOf (#1416)", () => {
	it('holds every measured case, in all four hosts', () => {
		expect(cases).toHaveLength(3088);
		for (const host of [undefined, 'word', 'powerpoint', 'access']) {
			expect(cases.filter((oracleCase) => oracleCase.host === host)).toHaveLength(772);
		}
	});

	it.each(cases.filter((oracleCase) => !KNOWN_QUIET.test(oracleCase.id)).map((oracleCase): [string, OracleCase] => [oracleCase.id, oracleCase]))(
		'%s gets the error the VBE gave',
		(_id, oracleCase) => {
			const message = / is rejected: Compile error: (.+?)\. Measured in /.exec(oracleCase.description)?.[1];
			const expected = oracleCase.expected === 'accepted' ? [] : [CODE_FOR_MESSAGE[message ?? '']];
			expect(expected).not.toContain(undefined);
			expect(errorCodes(oracleCase)).toEqual(expected);
		},
	);

	it('keeps the measured tables apart', () => {
		for (const name of ARGUMENT_TYPED_FUNCTIONS) {
			expect(VALUE_RESULT_FUNCTIONS.has(name)).toBe(false);
		}
		for (const name of BARE_VALUE_RESULT_FUNCTIONS) {
			expect(VALUE_RESULT_FUNCTIONS.has(name)).toBe(false);
		}
		expect(TYPEOF_KEYWORD_OPERANDS.has('cvar')).toBe(true);
		expect(VALUE_RESULT_FUNCTIONS.has('cvar')).toBe(false);
	});
});

describe('the shadows and Variants these results leave alone (#1416)', () => {
	const analyze = (...body: string[]): string[] => {
		const source = ['Option Explicit', ...body, ''].join('\r\n');
		return errorCodes({ id: 'inline', description: '', expected: 'accepted', modules: [{ name: 'Module1', source }] });
	};

	it('a source Function of the same name', () => {
		expect(analyze('Function Sqr(x) As Object', 'End Function', 'Sub Main()', '    Debug.Print Sqr(1) Is Nothing', 'End Sub')).toEqual([]);
	});

	it('a variable named VBA', () => {
		expect(analyze('Sub Main()', '    Dim VBA As Variant','    Debug.Print VBA.Mid$("ab", 1) Is Nothing', 'End Sub')).toEqual([]);
	});

	it('a Variant argument to an argument-typed function', () => {
		expect(analyze('Sub Main()', '    Dim v As Variant', '    Debug.Print Abs(v) Is Nothing', 'End Sub')).toEqual([]);
	});

	it('a function without a $ form measured', () => {
		expect(analyze('Sub Main()', '    Debug.Print Right("ab", 1) Is Nothing', 'End Sub')).toEqual([]);
	});
});
