import { readFileSync } from 'node:fs';
import { describe, expect, it, vi } from 'vitest';
import { diagnosticMetadataForCode } from '../../src/analyzer';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';
import { buildVbaProjectIndex, projectAnalysisOptionsForModule, projectProcedureSignatures, type VbaProjectModuleInput } from '../../src/vbaProjectAnalysis';
import * as expressions from '../../src/analyzer/parser/parseExpression';
import { forEachExpressionInBody, statementValueExpressions } from '../../src/analyzer/diagnostics/exprWalk';
import { parseModule } from '../../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../../src/analyzer/symbols/buildModuleSymbols';
import { expressionShapeResolver } from '../../src/analyzer/diagnostics/expressionShape';
import { resolveRuntimeFunction } from '../../src/analyzer/runtime/vbaRuntime';

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
function compileErrors(text: string, extraModules: VbaProjectModuleInput[] = [], moduleKind: 'standard' | 'class' | 'userform' = 'standard', implicitMembers: VbaProjectModuleInput['implicitMembers'] = []) {
	const project = buildVbaProjectIndex([{ moduleName: 'Module1', type: moduleKind, source: text, implicitMembers },
		{ moduleName: 'ThisWorkbook', type: 'document', documentType: 'workbook', source: '' },
		{ moduleName: 'Sheet1', type: 'document', documentType: 'worksheet', designerClass: 'Excel.Worksheet', implicitMembers: [], source: '' }, ...extraModules]);
	return analyzeVbaModuleSource({ source: text, moduleKind, moduleName: 'Module1', host: 'excel', referencedLibraries: ['VBA', 'Excel', 'stdole', 'Office'],
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
	it('reports a file Print array field operand once', () => {
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
				if (code === 'is-operator-non-object' && test.description.includes('Type mismatch')) {
					expect(errors.find((diag) => diag.code === code)?.message).toContain('VBE compile error: Type mismatch.');
				}
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
		if (['rq6_left_dollar_is_nothing', 'rq6_numbers_arg_plus', 'rq6_numbers_arg_is', 'rq6_numbers_arg_typeof'].includes(test.id)) {
			expect(errors[0].message).toContain('Type mismatch');
		}
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

function measuredMain(body: string, declarations = ''): string {
	return `Option Explicit\n${declarations}\nPublic Sub Main()\n${body}\nEnd Sub\n`;
}

describe('Excel 16.0 measured gap follow-ups', () => {
	it.each(['InStr("abc", "a")', 'Len("abc")'])(
		'keeps the measured scalar compile error even in an unexecuted branch: %s', (call) => {
			for (const prefix of ['', 'If False Then ']) {
				const errors = compileErrors(measuredMain(`${prefix}Debug.Print ${call} Is Nothing`));
				expect(errors.map((diag) => diag.code)).toEqual(['is-operator-non-object']);
				expect(errors[0].message).toContain('VBE compile error: Type mismatch.');
			}
		});
	it.each(['Year(#1/1/2026#)', 'VBA.Year(#1/1/2026#)', 'Left("abc", 1)', 'VBA.Left("abc", 1)', 'Flexible()'])(
		'accepts declared Variant results for Is: %s', (call) => {
			const declaration = 'Private Function Flexible() As Variant\nEnd Function\n';
			for (const expression of [`${call} Is Nothing`, `Nothing Is ${call}`]) {
				expect(compileErrors(measuredMain(`Debug.Print ${expression}`, declaration))).toEqual([]);
			}
		});
	it.each(['Print', 'Write'])('checks file %s values without treating the array target as a value', (statement) => {
		const errors = compileErrors(measuredMain(`Dim a(1) As Long\n${statement} #1, a + 1`));
		expect(errors.map((diag) => diag.code)).toEqual(['non-scalar-binary-operand']);
		expect(errors[0].message).toContain('Type mismatch');
		for (const expression of ['a(0) + 1', '1', 'a ^ 2', 'Not a']) {
			expect(compileErrors(measuredMain(`Dim a(1) As Long\n${statement} #1, ${expression}`))).toEqual([]);
		}
	});
	it('checks only ReDim bound expressions', () => {
		const errors = compileErrors(measuredMain('Dim a() As Long, n As Long\nReDim a(n Is Nothing)'));
		expect(errors.map((diag) => diag.code)).toEqual(['is-operator-non-object']);
		expect(errors[0].message).toContain('VBE compile error: Type mismatch.');
		for (const body of ['ReDim a(n + 1)', 'ReDim a(0 To n, 1 To n + 1)', 'ReDim Preserve a(n + 1)', 'ReDim a(Len("abc"))']) {
			expect(compileErrors(measuredMain(`Dim a() As Long, n As Long\n${body}`))).toEqual([]);
		}
	});
	it.each(['((Nothing)) + 1', '1 + ((Nothing))', '-((Nothing))', 'Not ((Nothing))'])(
		'unwraps nested Nothing operands: %s', (expression) => {
			const errors = compileErrors(measuredMain(`Debug.Print ${expression}`));
			expect(errors.map((diag) => diag.code)).toEqual(['non-scalar-binary-operand']);
			expect(errors[0].message).toContain('VBE compile error: Invalid use of object.');
		});
	it.each(['((1)) + 1', '((Nothing)) Is Nothing', 'Nothing Is ((Nothing))'])(
		'accepts nested parentheses controls: %s', (expression) => {
			expect(compileErrors(measuredMain(`Debug.Print ${expression}`))).toEqual([]);
		});
	const record = 'Private Type R\nn As Long\nEnd Type\n';
	it.each(['Print', 'Write'])('gives the shared visitor ownership of file %s UDT operands', (statement) => {
		const errors = compileErrors(measuredMain(`Dim r1 As R\n${statement} #1, r1 + 1`, record));
		expect(errors.map((diag) => diag.code)).toEqual(['non-scalar-binary-operand']);
		expect(errors[0].message).toContain('Type mismatch');
	});
	it.each(['Print', 'Write'])('accepts file %s scalar fields and the measured aggregate controls', (statement) => {
		for (const expression of ['r1.n + 1', 'r1 ^ 2', '2 ^ r1', 'Not r1']) {
			expect(compileErrors(measuredMain(`Dim r1 As R\n${statement} #1, ${expression}`, record)), expression).toEqual([]);
		}
	});
	it.each(['Debug.Print r1 + 1', 'If r1 + 1 Then Debug.Print 1', 'If True Then Debug.Print r1 + 1'])(
		'reports a walker-covered UDT operand only once: %s', (body) => {
			expect(compileErrors(measuredMain(`Dim r1 As R\n${body}`, record)).map((diag) => diag.code))
				.toEqual(['non-scalar-binary-operand']);
		});

	const enumeration = 'Private Enum E\na = 1\nEnd Enum\n';
	it.each([
		['Debug.Print E', 'Type mismatch'],
		['If E Then Debug.Print 1', 'Expected variable or procedure, not enum type'],
		['Dim x As Long\nx = E', 'Expected variable or procedure, not enum type'],
	])('matches the Enum value diagnosis for %s', (body, message) => {
		const errors = compileErrors(measuredMain(body, enumeration));
		expect(errors.map((diag) => diag.code)).toEqual(['malformed-statement']);
		expect(errors[0].message).toContain(`VBE compile error: ${message}.`);
	});
	it.each(['Debug.Print E.a', 'If E.a Then Debug.Print 1', 'Dim x As Long\nx = E.a'])(
		'accepts the Enum member control: %s', (body) => {
			expect(compileErrors(measuredMain(body, enumeration))).toEqual([]);
		});

	const countProperty = 'Public Property Get Count() As Long\nCount = 0\nEnd Property\nPublic Property Let Count(ByVal value As Long)\nEnd Property\n';
	it('selects Property Get for a read of a Get/Let pair', () => {
		expect(compileErrors(measuredMain('Debug.Print Count Is Nothing', countProperty), [], 'class').map((diag) => diag.code))
			.toEqual(['is-operator-non-object']);
	});
	it('accepts scalar property reads/writes and object/Variant property identity controls', () => {
		expect(compileErrors(measuredMain('Count = 1\nDebug.Print Count', countProperty), [], 'class')).toEqual([]);
		const objectProperty = 'Public Property Get Count() As Object\nSet Count = Nothing\nEnd Property\nPublic Property Set Count(ByVal value As Object)\nEnd Property\n';
		expect(compileErrors(measuredMain('Debug.Print Count Is Nothing', objectProperty), [], 'class')).toEqual([]);
		expect(compileErrors(measuredMain('Debug.Print Count Is Nothing', countProperty.replaceAll('As Long', 'As Variant')), [], 'class')).toEqual([]);
	});

	const numbers = 'Public Function Numbers() As Long()\nEnd Function\n';
	it('retains the array shape of a module-qualified function result', () => {
		expect(compileErrors(measuredMain('Debug.Print Module1.Numbers() + 1', numbers)).map((diag) => diag.code))
			.toEqual(['non-scalar-binary-operand']);
	});
	it('resolves the declared scalar result of VBA.CLng', () => {
		expect(compileErrors(measuredMain('Debug.Print VBA.CLng(1) Is Nothing')).map((diag) => diag.code))
			.toEqual(['is-operator-non-object']);
	});
	it.each([
		['Debug.Print Module1.Numbers()(0) Is Nothing', numbers],
		['Debug.Print Module1.Numbers() + 1', numbers.replace('As Long()', 'As Long')],
		['Debug.Print VBA.CLng(1) + 1', ''],
		['Debug.Print VBA.Left("x", 1) Is Nothing', ''],
		['Dim Module1 As Object\nDebug.Print Module1.Numbers() + 1', numbers],
		['Dim VBA As Object\nDebug.Print VBA.CLng(1) Is Nothing', ''],
	])('accepts qualified-call controls: %s', (body, declarations) => {
		expect(compileErrors(measuredMain(body, declarations))).toEqual([]);
	});
	it('leaves invalid qualified array invocations to argument-count', () => {
		expect(compileErrors(measuredMain('Debug.Print Module1.Numbers(1) + 1', numbers)).map((diag) => diag.code))
			.toEqual(['argument-count']);
	});
	it.each([10, 1000])('reuses declaration facts across qualified reads in a %i-declaration snapshot', (count) => {
		const declarations = numbers + Array.from({ length: count }, (_, i) => `Private Function F${i}() As Long\nEnd Function\n`).join('');
		const text = measuredMain('Debug.Print Module1.Numbers() + 1', declarations);
		const mod = parseModule(text);
		const symbols = buildModuleSymbols('Module1', 'standard', text, { parsedModule: mod });
		let reads = 0;
		for (const symbol of symbols.root.children ?? []) {
			const name = symbol.name;
			Object.defineProperty(symbol, 'name', { get() { reads++; return name; } });
		}
		const proc = mod.members.find((node) => node.kind === 'Procedure' && node.name === 'Main')!;
		if (proc.kind !== 'Procedure' || proc.body[0].kind !== 'Statement') { throw new Error('missing Main/Print'); }
		const expr = statementValueExpressions(text, proc.body[0])[0];
		if (expr.exprKind !== 'BinaryExpr') { throw new Error('missing operator'); }
		const factory = expressionShapeResolver(symbols);
		const shapeOf = factory(proc);
		expect(shapeOf(expr.left)?.isArray).toBe(true);
		const initialReads = reads;
		for (let i = 0; i < 100; i++) {
			expect(expressionShapeResolver(symbols)).toBe(factory);
			expect(factory(proc)).toBe(shapeOf);
			expect(shapeOf(expr.left)?.isArray).toBe(true);
		}
		expect(reads).toBe(initialReads);
	});

	it('parses the value after Case Is > before checking array operands', () => {
		expect(compileErrors(measuredMain('Dim a(1) As Long\nSelect Case 1\nCase Is > a + 1\nEnd Select')).map((diag) => diag.code))
			.toEqual(['non-scalar-binary-operand']);
	});
	it.each(['Case Is > a(0) + 1', 'Case Is > 1', 'Case 0, Is > a(0) + 1', 'Case 0 To a(0) + 1', 'Case Else'])(
		'accepts the Case control: %s', (clause) => {
			expect(compileErrors(measuredMain(`Dim a(1) As Long\nSelect Case 1\n${clause}\nEnd Select`))).toEqual([]);
		});
	it('caches Case clause parsing for the statement snapshot', () => {
		const text = measuredMain('Dim a(1) As Long\nSelect Case 1\nCase Is > a + 1\nEnd Select');
		const proc = parseModule(text).members.find((node) => node.kind === 'Procedure')!;
		if (proc.kind !== 'Procedure') { throw new Error('missing procedure'); }
		const block = proc.body.find((node) => node.kind === 'SelectBlock')!;
		if (block.kind !== 'SelectBlock') { throw new Error('missing Select'); }
		const stmt = block.body.find((node) => node.kind === 'Statement')!;
		if (stmt.kind !== 'Statement') { throw new Error('missing Case'); }
		const spy = vi.spyOn(expressions, 'parseExpression');
		try {
			const roots = statementValueExpressions(text, stmt);
			for (let i = 0; i < 100; i++) { expect(statementValueExpressions(text, stmt)).toBe(roots); }
			expect(roots.map((expr) => text.slice(expr.span.start, expr.span.end))).toEqual(['a + 1']);
			expect(spy).toHaveBeenCalledTimes(1);
		} finally { spy.mockRestore(); }
	});

	const needsValue = 'Function NeedsValue(ByVal arg As Long) As Long\nNeedsValue = arg\nDebug.Print NeedsValue\nIf NeedsValue Then Debug.Print 1\nEnd Function\n';
	it.each(['Debug.Print NeedsValue', 'If NeedsValue Then Debug.Print 1'])(
		'rejects a required-argument Function read in %s', (body) => {
			const errors = compileErrors(measuredMain(body, needsValue));
			expect(errors.map((diag) => diag.code)).toEqual(['argument-count']);
			expect(errors[0].message).toContain('Argument not optional');
		});
	it.each([
		['Debug.Print NeedsValue(1)', needsValue],
		['If NeedsValue(1) Then Debug.Print 1', needsValue],
		['Debug.Print NeedsValue\nIf NeedsValue Then Debug.Print 1', needsValue.replace('ByVal arg As Long', 'Optional ByVal arg As Long = 1')],
		['Dim NeedsValue As Long\nDebug.Print NeedsValue\nIf NeedsValue Then Debug.Print 1', needsValue],
	])('accepts Function value controls, including its own return variable: %s', (body, declarations) => {
		expect(compileErrors(measuredMain(body, declarations))).toEqual([]);
	});
	it('does not classify a bare Function call statement as a value read', () => {
		const errors = compileErrors(measuredMain('NeedsValue', needsValue));
		expect(errors.some((diag) => diag.message.includes('is read here with none'))).toBe(false);
	});

	it('keeps Left without $ out of compile errors for Is Nothing (runtime 424)', () => {
		expect(compileErrors(measuredMain('Debug.Print Left("x", 1) Is Nothing'))).toEqual([]);
	});
});

// Excel 16.0 measurements supplied for this branch; all use Option Explicit
// except the explicit DefLng implicit-local samples below.
describe('Excel 16.0 measured shape corrections', () => {
	it.each([
		'Year(d)', 'Month(d)', 'Day(d)', 'Weekday(d)', 'Hour(d)', 'Minute(d)', 'Second(d)',
		'DatePart("yyyy", d)', 'DateDiff("d", d, d)', 'Round(1.25, 1)', 'VBA.Year(d)',
	])('keeps the Variant compile shape of %s (Is runs with 424)', (call) => {
		const runtime = resolveRuntimeFunction(call.replace(/^VBA\./, '').split('(')[0]);
		expect(runtime?.declaredReturns).toBe('Variant');
		expect(runtime?.returns).toBe(call.startsWith('Round') ? 'Double' : call.startsWith('DateDiff') ? 'Long' : 'Integer');
		const declarations = 'Dim d As Date';
		expect(compileErrors(measuredMain(`${declarations}\nDebug.Print ${call} Is Nothing`))).toEqual([]);
		expect(compileErrors(measuredMain(`${declarations}\nIf TypeOf ${call} Is Collection Then Exit Sub`))).toEqual([]);
	});
	it.each([
		'Val("1")', 'Asc("a")', 'CInt(1)', 'Sgn(1)', 'StrComp("a", "b")',
		'Abs(1)', 'Int(1)', 'Fix(1)', 'Timer', 'CDate(1)', 'CStr(1)', 'CLng(1)',
		'Len("a")', 'InStr("a", "a")', 'Left$("a", 1)',
	])('preserves the measured scalar compile constraint of %s', (call) => {
		for (const [body, code] of [
			[`Debug.Print ${call} Is Nothing`, 'is-operator-non-object'],
			[`If TypeOf ${call} Is Collection Then Exit Sub`, 'typeof-invalid-expression'],
		]) {
			const errors = compileErrors(measuredMain(body));
			expect(errors.map((diag) => diag.code)).toEqual([code]);
			expect(errors[0].message).toContain('VBE compile error: Type mismatch.');
		}
		expect(compileErrors(measuredMain(`Debug.Print ${call}`))).toEqual([]);
	});
	it.each(['Case a', 'Case Is > a'])('rejects a whole array in %s', (clause) => {
		const errors = compileErrors(measuredMain(`Dim a(1) As Long\nSelect Case 1\n${clause}\nEnd Select`));
		expect(errors.map((diag) => diag.code)).toEqual(['non-scalar-binary-operand']);
		expect(errors[0].message).toContain('VBE compile error: Type mismatch.');
		for (const control of ['a(0)', '1']) {
			expect(compileErrors(measuredMain(`Dim a(1) As Long\nSelect Case 1\n${clause.replace(/a$/, control)}\nEnd Select`))).toEqual([]);
		}
	});
	it('rejects a whole array as the For start value', () => {
		const errors = compileErrors(measuredMain('Dim a(1) As Long, i As Long\nFor i = a To 1\nNext'));
		expect(errors.map((diag) => diag.code)).toEqual(['non-scalar-binary-operand']);
		expect(errors[0].message).toContain('VBE compile error: Type mismatch.');
		for (const start of ['a(0)', '1']) {
			expect(compileErrors(measuredMain(`Dim a(1) As Long, i As Long\nFor i = ${start} To 1\nNext`))).toEqual([]);
		}
	});
	it('types an undeclared read by DefLng without requiring an assignment', () => {
		const text = 'DefLng A-W\nSub Main()\nIf a Is Nothing Then Exit Sub\nEnd Sub\n';
		const errors = compileErrors(text);
		expect(errors.map((diag) => diag.code)).toEqual(['is-operator-non-object']);
		expect(errors[0].message).toContain("'a' has type Long");
		expect(errors[0].message).toContain('VBE compile error: Type mismatch.');
		for (const control of [text.replace('DefLng A-W\n', ''), text.replace('a Is Nothing', 'x Is Nothing'), text.replace('a Is Nothing', 'Application Is Nothing'), text.replace('a Is Nothing', 'Err Is Nothing'), text.replace('a Is Nothing', 'Now Is Nothing'), text.replace('Sub Main()\n', 'Sub Main()\nDim a As Object\n')]) {
			expect(compileErrors(control)).toEqual([]);
		}
		expect(compileErrors('Option Explicit\n' + text).map((diag) => diag.code)).toEqual(['undeclared-variable']);
	});
	it('rejects TypeOf on a variable of a same-module Enum', () => {
		const declaration = 'Private Enum SomeEnum\nfirst = 1\nEnd Enum\n';
		const body = 'Dim e As SomeEnum\nIf TypeOf e Is Collection Then Exit Sub';
		const errors = compileErrors(measuredMain(body, declaration));
		expect(errors.map((diag) => diag.code)).toEqual(['typeof-invalid-expression']);
		expect(errors[0].message).toBe('TypeOf requires an object expression. This is a VBE compile error: Type mismatch.');
		for (const type of ['Variant', 'Object', 'Collection']) {
			expect(compileErrors(measuredMain(body.replace('As SomeEnum', `As ${type}`), declaration))).toEqual([]);
		}
	});
	it.each(['Foo(1)', 'Module1.Foo(1)'])('leaves an invalid scalar invocation to argument-count: %s', (call) => {
		const declaration = 'Function Foo(n As Long, m As Long) As Long\nEnd Function\n';
		const errors = compileErrors(measuredMain(`If TypeOf ${call} Is Object Then Exit Sub`, declaration));
		expect(errors.map((diag) => diag.code)).toEqual(['argument-count']);
		expect(errors[0].message).toContain('Argument not optional');
		const valid = compileErrors(measuredMain(`If TypeOf ${call.replace('(1)', '(1, 2)')} Is Object Then Exit Sub`, declaration));
		expect(valid.map((diag) => diag.code)).toEqual(['typeof-invalid-expression']);
		expect(valid[0].message).toContain('VBE compile error: Type mismatch.');
		expect(compileErrors(measuredMain(`If TypeOf ${call.replace('(1)', '(1, 2)')} Is Object Then Exit Sub`, declaration.replace('As Long\n', 'As Object\n')))).toEqual([]);
	});
	it('names the scalar-variable Is diagnosis Type mismatch', () => {
		const errors = compileErrors(measuredMain('Dim n As Long\nDebug.Print n Is Nothing'));
		expect(errors.map((diag) => diag.code)).toEqual(['is-operator-non-object']);
		expect(errors[0].message).toBe("The 'Is' operator requires object operands, but 'n' has type Long, which is not an object. This is a VBE compile error: Type mismatch.");
	});
});

// Follow-up Excel measurements: use dead branches for function returns so
// argument/runtime failures cannot be mistaken for compile-time constraints.
describe('Excel-measured follow-up expression gaps', () => {
	it.each(['For i = 1 To a', 'For i = 1 To 2 Step a', '10  For i = a To 1'])(
		'rejects a bare array in %s', (header) => {
			const body = `Dim a(1) As Long, i As Long\n${header}\nNext`;
			const errors = compileErrors(measuredMain(body));
			expect(errors.map((d) => d.code)).toEqual(['non-scalar-binary-operand']);
			expect(errors[0].message).toBe("'a' is declared as an array, which For cannot read as one value. This is a VBE compile error: Type mismatch.");
			for (const control of ['a(0)', '1']) {
				expect(compileErrors(measuredMain(body.replace(header, header.replace(/\ba\b/, control))))).toEqual([]);
			}
		});
	it.each(['Case a, 1', 'Case 0, a', 'Case a To 1'])(
		'rejects a bare array in %s', (clause) => {
			const body = `Dim a(1) As Long\nSelect Case 1\n${clause}\nEnd Select`;
			const errors = compileErrors(measuredMain(body));
			expect(errors.map((d) => d.code)).toEqual(['non-scalar-binary-operand']);
			expect(errors[0].message).toBe("'a' is declared as an array, which Case cannot read as one value. This is a VBE compile error: Type mismatch.");
			for (const control of ['a(0)', '1']) {
				expect(compileErrors(measuredMain(body.replace(clause, clause.replace(/\ba\b/, control))))).toEqual([]);
			}
		});
	it('rejects a bare array as a ReDim bound', () => {
		const body = 'Dim a(1) As Long, d() As Long\nReDim d(a)';
		const errors = compileErrors(measuredMain(body));
		expect(errors.map((d) => d.code)).toEqual(['non-scalar-binary-operand']);
		expect(errors[0].message).toBe("'a' is declared as an array, which ReDim bound cannot read as one value. This is a VBE compile error: Type mismatch.");
		for (const control of ['a(0)', '1']) {
			expect(compileErrors(measuredMain(body.replace('d(a)', `d(${control})`)))).toEqual([]);
		}
	});
	it('visits the condition of an If nested in a single-line Else exactly once', () => {
		const body = 'Dim n As Long, s As String\nIf n Then Exit Sub Else If s Is Nothing Then Exit Sub';
		const errors = compileErrors(measuredMain(body));
		expect(errors.map((d) => d.code)).toEqual(['is-operator-non-object']);
		expect(errors[0].message).toBe("The 'Is' operator requires object operands, but 's' has type String, which is not an object. This is a VBE compile error: Type mismatch.");
		for (const type of ['Object', 'Variant']) {
			expect(compileErrors(measuredMain(body.replace('s As String', `s As ${type}`)))).toEqual([]);
		}
	});
	it.each([
		['-n', 'Dim n As Long', '', 'the expression has a scalar numeric type'],
		['(n + 1)', 'Dim n As Long', '', 'the expression has a scalar numeric type'],
		['S2', 'Const S2 = "abc"', '', "'S2' has type String"],
		['e1', '', 'Private Enum E\ne1 = 1\nEnd Enum\n', "'e1' has type E"],
		['n', 'Dim n As E', 'Private Enum E\ne1 = 1\nEnd Enum\n', "'n' has type E"],
		['Time$', '', '', "'Time$' has type String"],
	])('rejects the measured scalar Is operand %s', (operand, locals, declarations, detail) => {
		const errors = compileErrors(measuredMain(`${locals}\nDebug.Print ${operand} Is Nothing`, declarations));
		expect(errors.map((d) => d.code)).toEqual(['is-operator-non-object']);
		expect(errors[0].message).toBe(`The 'Is' operator requires object operands, but ${detail}, which is not an object. This is a VBE compile error: Type mismatch.`);
		expect(compileErrors(measuredMain(`${locals}\nDebug.Print ${operand}`, declarations))).toEqual([]);
	});
	it.each(['n', '-n', '(n + 1)'])('keeps unknown arithmetic inputs conservative: %s', (operand) => {
		expect(compileErrors(measuredMain(`Dim n As Variant\nDebug.Print ${operand} Is Nothing`))).toEqual([]);
	});
	it('respects local shadows of Enum members and string constants', () => {
		const declarations = 'Private Enum E\ne1 = 1\nEnd Enum\nConst S2 = "abc"\n';
		expect(compileErrors(measuredMain('Dim e1 As Object, S2 As Variant\nDebug.Print e1 Is Nothing, S2 Is Nothing', declarations))).toEqual([]);
	});
	it.each([
		['Dir', "'Dir' has type String"], ['MsgBox("x")', 'the expression has a scalar numeric type'],
		['VarType(1)', 'the expression has a scalar numeric type'], ['Rnd', "'Rnd' has type Single"],
		['FreeFile', "'FreeFile' has type Integer"], ['Replace("a", "a", "b")', 'the expression has type String'],
		['UBound(a)', 'the expression has type Long'], ['TypeName(1)', 'the expression has type String'],
		['IsNumeric(1)', 'the expression has type Boolean'], ['Sqr(1)', 'the expression has type Double'],
		['Join(a)', 'the expression has type String'], ['Val("1")', 'the expression has type Double'],
		['Asc("a")', 'the expression has type Integer'], ['CInt(1)', 'the expression has type Integer'],
		['CLng(1)', 'the expression has type Long'], ['Sgn(1)', 'the expression has type Integer'],
		['StrComp("a", "b")', 'the expression has type Integer'], ['Abs(1)', 'the expression has a scalar numeric type'],
		['Int(1)', 'the expression has a scalar numeric type'], ['Fix(1)', 'the expression has a scalar numeric type'],
		['Timer', "'Timer' has type Single"], ['CDate(1)', 'the expression has type Date'],
		['CStr(1)', 'the expression has type String'], ['Len("a")', 'the expression has type Long'],
		['InStr("a", "a")', 'the expression has type Long'],
	])('rejects the measured typed return %s in a dead branch', (call, detail) => {
		const body = `Dim a(1) As String\nIf False Then Debug.Print ${call} Is Nothing`;
		const errors = compileErrors(measuredMain(body));
		expect(errors.map((d) => d.code)).toEqual(['is-operator-non-object']);
		expect(errors[0].message).toBe(`The 'Is' operator requires object operands, but ${detail}, which is not an object. This is a VBE compile error: Type mismatch.`);
		expect(compileErrors(measuredMain(body.replace(' Is Nothing', '')))).toEqual([]);
	});
	it.each([
		'DateSerial(2026, 1, 1)', 'TimeSerial(1, 2, 3)', 'DateValue("2026-01-01")', 'TimeValue("01:02:03")',
		'DateAdd("d", 1, d)', 'Environ("PATH")', 'Format(1)', 'Chr(65)', 'Time', 'CDec(1)', 'Date', 'Now',
		'Hex(1)', 'Str(1)', 'Trim("x")', 'UCase("x")', 'Split("a,b", ",")',
		'Year(d)', 'Month(d)', 'Day(d)', 'Weekday(d)', 'Hour(d)', 'Minute(d)', 'Second(d)',
		'DatePart("yyyy", d)', 'DateDiff("d", d, d)', 'Round(1.25, 1)', 'Left("x", 1)',
	])('accepts the measured Variant return %s in a dead branch', (call) => {
		expect(compileErrors(measuredMain(`Dim d As Date\nIf False Then Debug.Print ${call} Is Nothing`))).toEqual([]);
		// Date is deliberately absent from runtime metadata to avoid a type-name collision.
		if (call !== 'Date') { expect(resolveRuntimeFunction(call.split('(')[0])?.declaredReturns).toBe('Variant'); }
	});
	it.each([
		['If Sheet1 Is Nothing Then Exit Sub', 'standard'],
		['If TypeOf Sheet1 Is Worksheet Then Exit Sub', 'standard'],
		['If Me Is Nothing Then Exit Sub', 'class'],
	] as const)('does not turn a project object into an implicit Long: %s', (body, kind) => {
		expect(compileErrors(`DefLng A-Z\nSub Main()\n${body}\nEnd Sub\n`, [], kind)).toEqual([]);
	});
	it('does not turn a predeclared class or UserForm into an implicit Long', () => {
		const modules = [
			{ moduleName: 'Factory', type: 'class', predeclaredId: true, source: '' },
			{ moduleName: 'UserForm1', type: 'userform', source: '' },
		];
		for (const name of ['Factory', 'UserForm1']) {
			expect(compileErrors(`DefLng A-Z\nSub Main()\nIf ${name} Is Nothing Then Exit Sub\nEnd Sub\n`, modules)).toEqual([]);
		}
	});
	it('excludes UserForm designer controls and implicit members from DefType fallback', () => {
		const controls = [{ name: 'CommandButton1', type: 'MSForms.CommandButton' }, { name: 'Picker', type: 'MSForms.ComboBox' }];
		for (const { name } of controls) {
			expect(compileErrors(`DefLng A-Z\nSub Main()\nIf ${name} Is Nothing Then Exit Sub\nEnd Sub\n`, [], 'userform', controls)).toEqual([]);
		}
	});
	it.each(['standardModule', 'class', 'document', 'userform'] as const)('excludes every project module kind from the implicit-local fallback: %s', (kind) => {
		const text = 'DefLng A-Z\nSub Main()\nDebug.Print OtherModule Is Nothing\nEnd Sub\n';
		const mod = parseModule(text);
		const symbols = buildModuleSymbols('Module1', 'standard', text, { parsedModule: mod });
		const proc = mod.members.find((m): m is import('../../src/analyzer/parser/nodes').ProcedureNode => m.kind === 'Procedure')!;
		const stmt = proc.body[0];
		if (stmt.kind !== 'Statement') { throw new Error('missing print'); }
		const expr = statementValueExpressions(text, stmt)[0];
		if (expr.exprKind !== 'BinaryExpr') { throw new Error('missing Is'); }
		const shape = expressionShapeResolver(symbols, undefined, { projectClassMembers: [{ name: 'OtherModule', moduleName: 'OtherModule', kind, members: [] }] })(proc);
		expect(shape(expr.left)).toBeUndefined();
	});
});
