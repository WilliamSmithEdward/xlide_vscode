import { readFileSync } from 'node:fs';
import { describe, expect, it, vi } from 'vitest';
import { diagnosticMetadataForCode } from '../../src/analyzer';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';
import { buildVbaProjectIndex, projectAnalysisOptionsForModule, projectProcedureSignatures, type VbaProjectModuleInput } from '../../src/vbaProjectAnalysis';
import * as expressions from '../../src/analyzer/parser/parseExpression';
import { forEachExpressionInBody, statementValueExpressions } from '../../src/analyzer/diagnostics/exprWalk';
import { parseModule } from '../../src/analyzer/parser/parseModule';

const contexts = [
	'result = {expr}', 'Debug.Print {expr}', 'Debug.Print 1, {expr}',
	'If {expr} Then\nEnd If', 'If False Then\nElseIf {expr} Then\nEnd If', 'If {expr} Then Exit Function',
	'Take {expr}', 'Call Take({expr})', 'result = Echo({expr})', 'result = CStr({expr})',
	'Select Case {expr}\nCase Else\nEnd Select', 'With holder\nTake {expr}\nEnd With',
	'Do While {expr}\nExit Do\nLoop', 'Do Until {expr}\nExit Do\nLoop',
	'Do\nExit Do\nLoop While {expr}', 'Do\nExit Do\nLoop Until {expr}', 'While {expr}\nWend',
	'For i = {expr} To 1\nNext', 'For i = 1 To {expr}\nNext', 'For i = 1 To 2 Step {expr}\nNext',
	'items(0) = {expr}', 'Take value:={expr}', 'result = "x" & ({expr})',
	'holder.Item(1) = {expr}', 'Main = {expr}',
];
const declarations = 'Private Type Record\nfield As Long\nEnd Type\nPrivate Enum Choice\none = 1\nEnd Enum\n';
function source(body: string, extra = ''): string {
	return declarations + extra + '\nFunction Main() As Variant\nDim n As Long, result As Variant, i As Long\nDim a(1) As Long, items(2) As Variant\nDim rec As Record, holder As Object, flexible As Variant\nConst limit = 2\n' + body + '\nEnd Function\nSub Action()\nEnd Sub\nSub Take(ByVal value As Variant)\nEnd Sub\nFunction Echo(ByVal value As Variant) As Variant\nEnd Function\nFunction NumberValue() As Long\nEnd Function\n';
}
function compileErrors(text: string, extraModules: VbaProjectModuleInput[] = []) {
	const project = buildVbaProjectIndex([{ moduleName: 'Module1', type: 'standard', source: text },
		{ moduleName: 'ThisWorkbook', type: 'document', documentType: 'workbook', source: '' },
		{ moduleName: 'Sheet1', type: 'document', documentType: 'worksheet', designerClass: 'Excel.Worksheet', implicitMembers: [], source: '' }, ...extraModules]);
	return analyzeVbaModuleSource({ source: text, moduleKind: 'standard', moduleName: 'Module1', host: 'excel', referencedLibraries: ['VBA', 'Excel', 'stdole', 'Office'],
		workbookSheets: [{ name: 'Sheet1', kind: 'worksheet' }],
		...projectAnalysisOptionsForModule(project, 'Module1', projectProcedureSignatures(project)),
	}).diagnostics.filter((diag) => diag.severity === 'error' && diagnosticMetadataForCode(diag.code ?? '')?.diagnosticKind === 'compile-error');
}

// Names deliberately differ from the measured fixtures: these are statement
// grammar and declaration-shape checks, independent of a particular callee.
describe('the same expression rules in every statement context', () => {
	for (const [expression, code] of [
		['n.Value', 'scalar-member-access'], ['n(1)', 'scalar-indexed'],
		['limit(1)', 'scalar-indexed'], ['rec.missing', 'member-not-found'],
		['Choice.missing', 'member-not-found'], ['Choice', 'malformed-statement'], ['Me', 'me-outside-object-module'],
		['Action', 'sub-used-as-value'], ['Action()', 'sub-used-as-value'],
		['TypeOf n Is Collection', 'typeof-invalid-expression'],
		['n Is Nothing', 'is-operator-non-object'], ['1 Is 1', 'is-operator-non-object'],
		['a + 1', 'non-scalar-binary-operand'], ['((a)) & "x"', 'non-scalar-binary-operand'],
	] as const) {
		it.each(contexts)(`${expression} in %s`, (context) => {
			const text = source(context.replace('{expr}', expression));
			const hits = compileErrors(text).filter((diag) => diag.code === code);
			expect(hits, compileErrors(text).map((diag) => `${diag.code}: ${diag.message}`).join('\n')).toHaveLength(1);
		});
	}
	it.each(contexts)('accepts flexible values in %s', (context) => {
		expect(compileErrors(source(context.replace('{expr}', 'flexible')))).toEqual([]);
	});
	it.each(contexts)('accepts object TypeOf in %s', (context) => {
		expect(compileErrors(source(context.replace('{expr}', 'TypeOf holder Is Collection')))).toEqual([]);
	});
	it.each(['(n)', '((n))', 'NumberValue()', 'a(0)', 'CLng(1)', 'rec.field'])('resolves scalar Is operands: %s', (operand) => {
		for (const expression of [`${operand} Is Nothing`, `Nothing Is ${operand}`]) {
			expect(compileErrors(source(`Debug.Print ${expression}`)).filter((diag) => diag.code === 'is-operator-non-object')).toHaveLength(1);
		}
	});
	it.each(['Let Set c = Nothing', 'If True Then Let Set c = Nothing', 'again: Let Set c = Nothing', '10 Let Set c = Nothing'])(
		'rejects the malformed assignment prefix exactly once: %s', (body) => {
			const errors = compileErrors(source(`Dim c As Collection\n${body}`));
			expect(errors.map((diag) => diag.code)).toEqual(['malformed-statement']);
			expect(errors[0].message).toContain('Syntax error');
		});
	it.each(['Set c = Nothing', 'Set c = New Collection', 'Let n = 1', 'If c Is Nothing Then Set c = New Collection'])(
		'accepts valid assignment syntax: %s', (body) => {
			expect(compileErrors(source(`Dim c As Collection\n${body}`))).toEqual([]);
		});
	it('leaves arguments to an array-returning function to argument-count', () => {
		const errors = compileErrors(source('Debug.Print ArrayValue(0) + 1', 'Function ArrayValue() As Long()\nEnd Function\n'));
		expect(errors.map((diag) => diag.code)).toEqual(['argument-count']);
	});
	it.each([
		['result = TypeOf holder Is Choice', []],
		['Debug.Print TypeOf holder Is Choice', []],
		['Set holder = New Choice', ['invalid-new-type-name']],
	] as const)(
		'does not describe a type name as an Enum value: %s', (body, codes) => {
			expect(compileErrors(source(body)).map((diag) => diag.code)).toEqual(codes);
		});
	it('accepts a TypeOf class name shared with a Sub', () => {
		const extra = [{ moduleName: 'Action', type: 'class', source: '' }];
		expect(compileErrors(source('Dim c As Object\nIf TypeOf c Is Action Then\nEnd If'), extra)).toEqual([]);
	});
	it('leaves indexing an array call result to runtime', () => {
		const text = source('Debug.Print Numbers()(0) Is Nothing', 'Function Numbers() As Long()\nEnd Function\n');
		expect(compileErrors(text)).toEqual([]);
	});
	it.each([
		['Numbers() + 1', 'non-scalar-binary-operand'],
		['TypeOf Numbers() Is Collection', 'typeof-invalid-expression'],
	] as const)('rejects a whole-array function result: %s', (expression, code) => {
		expect(compileErrors(source(`Debug.Print ${expression}`, 'Function Numbers() As Long()\nEnd Function\n')).map((diag) => diag.code))
			.toEqual([code]);
	});
	it.each([
		['Left$("x", 1) Is Nothing', 'the expression has type String'],
		['NumberValue() Is Nothing', 'the expression has type Long'],
		['ArrayValue() + 1', 'the expression has array type'],
	] as const)('describes expression types without claiming a declaration: %s', (expression, detail) => {
		const errors = compileErrors(source(`Debug.Print ${expression}`, 'Function ArrayValue() As Long()\nEnd Function\n'));
		expect(errors).toHaveLength(1);
		expect(errors[0].message).toContain(detail);
	});
	it('keeps actual Sub calls, AddressOf and local shadowing legal', () => {
		expect(compileErrors(source('Action\nCall Action()\nTake 1'))).toEqual([]);
		expect(compileErrors(source('Dim Action As Variant\nDebug.Print Action'))).toEqual([]);
		expect(compileErrors(source('Dim limit As Variant\nDebug.Print limit(1)').replace('Const limit = 2\n', ''))).toEqual([]);
	});
	it.each(['Action', 'Choice'])('does not read bang member %s as a bare value', (name) => {
		expect(compileErrors(source(`Debug.Print holder!${name}`))).toEqual([]);
		expect(compileErrors(source(`Debug.Print holder!${name}, ${name}`)).map((diag) => diag.code))
			.toEqual([name === 'Action' ? 'sub-used-as-value' : 'malformed-statement']);
	});
	it.each(['Action', 'Choice'])('does not read TypeOf type name %s as a bare value', (name) => {
		const errors = compileErrors(source(`Debug.Print TypeOf holder Is ${name}`));
		expect(errors.filter((diag) => ['sub-used-as-value', 'malformed-statement'].includes(diag.code ?? ''))).toHaveLength(0);
	});
	it.each(['n Is Nothing', 'Nothing Is n', '(n) Is Nothing'])('keeps the DefType of implicit locals in %s', (expression) => {
		const text = `DefLng N\nSub Main()\nn = 1\nDebug.Print ${expression}\nEnd Sub\n`;
		expect(compileErrors(text).map((diag) => diag.code)).toEqual(['is-operator-non-object']);
	});
	it.each(['Property', 'Action + Property'])('preserves Sub diagnostics when parsing fails on %s', (expression) => {
		const text = source(`result = ${expression}`, 'Sub Property()\nEnd Sub\n');
		expect(compileErrors(text).filter((diag) => diag.code === 'sub-used-as-value')).toHaveLength(expression === 'Property' ? 1 : 2);
	});
	it('preserves Enum diagnostics when parsing a contextual identifier fails', () => {
		const text = source('result = Property\nresult = TypeName(Property)', 'Private Enum Property\nitem = 1\nEnd Enum\n');
		expect(compileErrors(text).filter((diag) => diag.code === 'malformed-statement')).toHaveLength(2);
	});
	it.each(contexts)('gives TypeOf ownership of scalar fields in %s', (context) => {
		expect(compileErrors(source(context.replace('{expr}', 'TypeOf rec.field Is Collection'))).map((diag) => diag.code))
			.toEqual(['typeof-invalid-expression']);
	});
	it('checks values used to address assignment targets', () => {
		expect(compileErrors(source('items(Action) = 1')).filter((diag) => diag.code === 'sub-used-as-value')).toHaveLength(1);
		expect(compileErrors(source('items(0) = 1\nAction'))).toEqual([]);
	});
	it('preserves the Type-field fallback in statement grammars without expression ASTs', () => {
		const text = source('Dim row As Payload\nPrint #1, row.values + 1', 'Private Type Payload\nvalues(1) As Long\nEnd Type\n');
		expect(compileErrors(text).filter((diag) => diag.code === 'non-scalar-binary-operand')).toHaveLength(1);
	});
	it('keeps module-qualified Sub calls and implicit With arguments legal', () => {
		const extra = [{ moduleName: 'Foo', type: 'standard', source: 'Public Sub Foo()\nEnd Sub\n' }];
		expect(compileErrors(source('Foo.Foo\nCall Foo.Foo()'), extra)).toEqual([]);
		expect(compileErrors(source('With rec\nTake .field\nEnd With'))).toEqual([]);
	});
	it('respects a project procedure shadowing an intrinsic scalar function', () => {
		const extra = [{ moduleName: 'Other', type: 'standard', source: 'Public Function Len(ByVal text As String) As Object\nEnd Function\n' }];
		for (const expression of ['Len("x") Is Nothing', 'TypeOf Len("x") Is Collection']) {
			expect(compileErrors(source(`Debug.Print ${expression}`), extra)).toEqual([]);
		}
	});
	it('skips inactive expressions and visits each header exactly once', () => {
		const text = source('#If False Then\nDebug.Print n Is Nothing\n#End If\nIf n Is Nothing Then\nElseIf n Is Nothing Then\nEnd If\nDo\nLoop Until n Is Nothing\nFor i = 1 To a + 1 Step a + 2\nNext');
		const errors = compileErrors(text);
		expect(errors.filter((diag) => diag.code === 'is-operator-non-object')).toHaveLength(3);
		expect(errors.filter((diag) => diag.code === 'non-scalar-binary-operand')).toHaveLength(2);
	});
	it('shares fallback expression parsing across consumers of one statement snapshot', () => {
		const text = source('Debug.Print n Is Nothing, a + 1');
		const mod = parseModule(text);
		const proc = mod.members.find((node) => node.kind === 'Procedure')!;
		if (proc.kind !== 'Procedure') { throw new Error('missing procedure'); }
		const stmt = proc.body.find((node) => node.kind === 'Statement' && node.raw.startsWith('Debug.Print'))!;
		if (stmt.kind !== 'Statement') { throw new Error('missing statement'); }
		const spy = vi.spyOn(expressions, 'parseExpression');
		try {
			const first = statementValueExpressions(text, stmt);
			for (let i = 0; i < 100; i++) { expect(statementValueExpressions(text, stmt)).toBe(first); }
			expect(first).toHaveLength(2);
			expect(spy).toHaveBeenCalledTimes(2);
		} finally { spy.mockRestore(); }
	});
	it('does linear shared work for nested blocks, branches and print items', () => {
		for (const count of [10, 100]) {
			const text = source(Array(count).fill('If flexible Then\nElseIf n Is Nothing Then\nDebug.Print n Is Nothing, n Is Nothing\nEnd If\nDo\nLoop Until n Is Nothing').join('\n'));
			const mod = parseModule(text);
			const proc = mod.members.find((node) => node.kind === 'Procedure')!;
			if (proc.kind !== 'Procedure') { throw new Error('missing procedure'); }
			const offsets: number[] = [];
			forEachExpressionInBody(proc.body, undefined, (expr) => {
				if (expr.exprKind === 'BinaryExpr' && expr.operator === 'Is') { offsets.push(expr.span.start); }
			}, text);
			expect(offsets).toHaveLength(count * 4);
			expect(new Set(offsets).size).toBe(offsets.length);
		}
	});
});

const measured = (JSON.parse(readFileSync('syntax_corpus/oracle/vbe_oracle_cases.json', 'utf8')) as {
	cases: { id: string; description: string; source?: string; modules?: { name: string; source: string }[]; expected: string }[];
}).cases.filter((test) => /^(cl_ctx_|codex_expr_|cl_stmt_|cl_decl_enum_name_as_value)/.test(test.id)
	&& test.description.endsWith('(measured in Excel 16.0, Office Home 2024, 64-bit, 2026-10-15)'));

const audit = JSON.parse(readFileSync('syntax_corpus/diagnostic_influence_audit.json', 'utf8')) as {
	diagnostics: { code: string; diagnosticKind: string; assertedOracleCases: string[] }[];
};
const expectedFamilies = new Map<string, string[]>();
for (const rule of audit.diagnostics) {
	if (rule.diagnosticKind !== 'compile-error') { continue; }
	for (const id of rule.assertedOracleCases) {
		const families = expectedFamilies.get(id) ?? [];
		families.push(rule.code);
		expectedFamilies.set(id, families);
	}
}

describe('measured expression-context regressions', () => {
	it.each(measured)('$id ($expected)', (test) => {
		const errors = compileErrors(test.source ?? test.modules?.find((mod) => mod.name === 'Module1')?.source ?? '');
		expect(errors.length > 0).toBe(test.expected === 'rejected');
		if (test.expected === 'rejected') {
			for (const code of expectedFamilies.get(test.id) ?? []) {
				expect(errors.map((diag) => diag.code)).toContain(code);
			}
		}
	});
});

const round2Codes: Record<string, string[]> = {
	rq6_left_is_nothing: [],
	rq6_left_dollar_is_nothing: ['is-operator-non-object'],
	rq6_now_is_nothing: [],
	rq6_typeof_left: [],
	rq6_typeof_paren_collection_if: ['typeof-invalid-expression'],
	rq6_typeof_paren_collection_assign: ['typeof-invalid-expression'],
	rq6_numbers_arg_plus: ['non-scalar-binary-operand'],
	rq6_numbers_arg_is: ['is-operator-non-object'],
	rq6_numbers_arg_typeof: ['typeof-invalid-expression'],
	rq6_udt_pow_private: [],
	rq6_udt_pow_right_private: [],
};
const round2 = (JSON.parse(readFileSync('syntax_corpus/oracle/vbe_oracle_cases.json', 'utf8')) as {
	cases: { id: string; source: string; expected: string }[];
}).cases.filter((test) => test.id.startsWith('rq6_'));

describe('Excel 16.0 round-2 measurements (2026-10-16)', () => {
	it('keeps all eleven exact measured sources and audit links', () => {
		expect(round2.map((test) => test.id).sort()).toEqual(Object.keys(round2Codes).sort());
		for (const test of round2) {
			expect(test.expected).toBe(round2Codes[test.id].length ? 'rejected' : 'accepted');
			expect(expectedFamilies.has(test.id), test.id).toBe(true);
		}
	});
	it.each(round2)('$id reports exactly the measured compile diagnosis', (test) => {
		const errors = compileErrors(test.source);
		expect(errors.map((diag) => diag.code)).toEqual(round2Codes[test.id]);
		if (test.id.includes('paren_collection')) { expect(errors[0].message).toContain('Syntax error'); }
		if (test.id === 'rq6_left_dollar_is_nothing') { expect(errors[0].message).toContain('type String'); }
		if (test.id === 'rq6_numbers_arg_is') { expect(errors[0].message).toContain('array type'); }
	});
	it.each(['Right("x", 1)', 'Mid("x", 1)', 'Trim("x")', 'Space(1)', 'Hex(1)', 'Oct(1)', 'Str(1)', 'Now', 'Time', 'DateSerial(2026, 10, 16)'])(
		'uses Variant declarations for compile constraints: %s', (call) => {
			for (const expression of [`${call} Is Nothing`, `TypeOf ${call} Is Collection`]) {
				expect(compileErrors(source(`Debug.Print ${expression}`))).toEqual([]);
			}
		});
	it.each(['Right$("x", 1)', 'Mid$("x", 1)', 'Trim$("x")', 'Space$(1)', 'Hex$(1)', 'Oct$(1)', 'Str$(1)', 'Time$'])(
		'keeps typed $ function constraints: %s', (call) => {
			expect(compileErrors(source(`Debug.Print ${call} Is Nothing`)).map((diag) => diag.code)).toEqual(['is-operator-non-object']);
			expect(compileErrors(source(`Debug.Print TypeOf ${call} Is Collection`)).map((diag) => diag.code)).toEqual(['typeof-invalid-expression']);
		});
	it.each(['Numbers()', 'Numbers(1, 2)', 'Numbers(other:=1)'])(
		'leaves invalid array invocations to argument-count: %s', (call) => {
			for (const expression of [`${call} + 1`, `${call} Is Nothing`, `TypeOf ${call} Is Collection`]) {
				expect(compileErrors(source(`Debug.Print ${expression}`, 'Private Function Numbers(ByVal n As Long) As Long()\nEnd Function\n'))
					.map((diag) => diag.code)).toEqual(['argument-count']);
			}
		});
	it.each(['Numbers(1)', 'Numbers(n:=1)', 'Numbers()', 'Numbers(1, 2)'])(
		'preserves valid required, named, optional and ParamArray invocations: %s', (call) => {
			const declaration = call === 'Numbers()' ? 'Optional ByVal n As Long = 1'
				: call === 'Numbers(1, 2)' ? 'ParamArray n() As Variant' : 'ByVal n As Long';
			for (const [expression, code] of [[`${call} + 1`, 'non-scalar-binary-operand'],
				[`${call} Is Nothing`, 'is-operator-non-object'], [`TypeOf ${call} Is Collection`, 'typeof-invalid-expression']]) {
				expect(compileErrors(source(`Debug.Print ${expression}`, `Private Function Numbers(${declaration}) As Long()\nEnd Function\n`))
					.map((diag) => diag.code)).toEqual([code]);
			}
		});
	it.each(['TypeOf ((c)) Is Collection', 'TypeOf (((c))) Is Collection'])(
		'gives TypeOf ownership of nested Collection parentheses: %s', (expression) => {
			expect(compileErrors(source(`Dim c As New Collection\nDebug.Print ${expression}`)).map((diag) => diag.code))
				.toEqual(['typeof-invalid-expression']);
		});
});
