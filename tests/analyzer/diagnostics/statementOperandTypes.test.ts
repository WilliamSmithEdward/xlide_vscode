// A statement handed a variable of a type it refuses: Line Input # and the
// Mid statement into anything but a String or a Variant, Load and Unload on a
// literal, a scalar, a user-defined type or a whole array, Get and Put on an
// object or a user-defined type holding one. Measured in Excel, Word,
// PowerPoint and Access 16.0 (build 20527) through the VBE oracle; one form
// of each claim is a statement_operand_type_* fixture in
// syntax_corpus/oracle/vbe_oracle_cases.json, and the rows below follow the
// others measured the same way.

import * as fs from 'node:fs';
import * as path from 'node:path';
import { describe, expect, it } from 'vitest';
import { analyzeModule } from '../../../src/analyzer';
import { analyzeVbaModuleSource } from '../../../src/vbaModuleAnalysis';
import {
	buildVbaProjectIndex,
	effectiveModuleKind,
	projectAnalysisOptionsForModule,
	projectProcedureSignatures,
	type VbaProjectModuleInput,
} from '../../../src/vbaProjectAnalysis';
import { byCode, expectDiagnostic } from '../../helpers/diagnostics';

interface OracleCase {
	id: string;
	expected: 'accepted' | 'rejected';
	host: 'excel' | 'word' | 'powerpoint' | 'access';
	modules: { name: string; type: string; source: string }[];
}

const fixtures: OracleCase[] = (JSON.parse(
	fs.readFileSync(path.join(process.cwd(), 'syntax_corpus', 'oracle', 'vbe_oracle_cases.json'), 'utf8'),
) as { cases: OracleCase[] }).cases.filter((kase) => kase.id.startsWith('statement_operand_type_'));

/** Each error's code, each module analyzed with its project, as the extension does. */
function projectErrors(modules: VbaProjectModuleInput[], host: OracleCase['host']): string[] {
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
		.map((diagnostic) => diagnostic.code ?? ''));
}

const operandTypes = (source: string) => byCode(analyzeModule(source, { host: 'excel' }), 'statement-operand-type');

function procedure(dims: string, ...lines: string[]): string {
	return ['Option Explicit', TYPES, 'Sub Main()', `    Dim f As Integer${dims ? ', ' + dims : ''}`, '    f = FreeFile', ...lines.map((line) => '    ' + line), 'End Sub', ''].join('\r\n');
}

const TYPES = [
	'Private Type T1', '    s As String', '    n As Long', 'End Type',
	'Private Type T2', '    o As Object', 'End Type',
	'Private Type T3', '    t As T2', 'End Type',
	'Private Type T4', '    v As Variant', 'End Type',
	'Private Type T6', '    oa(2) As Object', 'End Type',
	'Private Type T7', '    oa(2) As Object', '    sa(2) As String', 'End Type',
].join('\r\n');

describe('a statement operand of a type the statement refuses', () => {
	it('covers every host', () => {
		expect(new Set(fixtures.map((kase) => kase.host))).toEqual(new Set(['excel', 'word', 'powerpoint', 'access']));
	});

	it.each(fixtures.map((kase) => [kase.id, kase] as const))('%s', (id, kase) => {
		const modules = kase.modules.map((mod) => ({ moduleName: mod.name, type: mod.type, source: mod.source }));
		const errors = projectErrors(modules, kase.host);
		// The Type as Mid's value is the existing udt-value-mismatch's.
		const code = id.includes('_mid_value_udt_') ? 'udt-value-mismatch' : 'statement-operand-type';
		expect(errors).toEqual(kase.expected === 'rejected' ? [code] : []);
	});

	it.each([
		['n As Long', 'Line Input #f, n', 'n', "Line Input # reads a line into a String or a Variant, and 'n' is declared As Long. This is a VBE compile error: Type mismatch."],
		['sa(2) As String', 'Line Input #f, sa', 'sa', "and 'sa' is an array."],
		['sa(2) As String', 'Line Input #f, sa()', 'sa()', "and 'sa()' is an array."],
		['fsa(2) As String * 3', 'Mid(fsa, 1) = "x"', 'fsa', "and 'fsa' is an array."],
		['ta(2) As T1', 'Mid(ta(0), 1) = "x"', 'ta(0)', "and 'ta(0)' is declared As T1."],
		['n As Long', 'Unload (n)', '(n)', "and '(n)' is declared As Long."],
		['oa(2) As Object', 'Unload oa', 'oa', "Unload takes an object, such as a UserForm, and 'oa' is an array."],
		['va() As Variant', 'Unload va', 'va', "and 'va' is an array."],
		['oa(2) As Object', 'Load oa', 'oa', "Load takes an object, such as a UserForm, and 'oa' is an array."],
		['n As Long, Mid(2, 2) As Long', 'Mid(n, 1) = 5', 'n', "The Mid statement writes into a String or a Variant, and 'n' is declared As Long."],
		['ba() As Byte', 'Mid(ba, 1) = "x"', 'ba', "and 'ba' is an array."],
		['t7 As T7', 'Mid(t7.sa(), 1) = "x"', 't7.sa()', "and 't7.sa()' is an array."],
		['t7 As T7', 'Put #f, , t7.oa(1)', 't7.oa(1)', "and 't7.oa(1)' is declared As Object."],
		['t As T1', 'Line Input #f, t.n', 't.n', "and 't.n' is declared As Long."],
		['o As Object', 'Line Input #f, o', 'o', "and 'o' is declared As Object."],
		['n As Long', 'Mid(n, 1) = "x"', 'n', "The Mid statement writes into a String or a Variant, and 'n' is declared As Long. This is a VBE compile error: Type mismatch."],
		['na(2) As Long', 'MidB$(na(1), 1, 1) = "x"', 'na(1)', "and 'na(1)' is declared As Long."],
		['', 'Unload "x"', '"x"', `Unload takes an object, such as a UserForm, and '"x"' is a literal. This is a VBE compile error: Type mismatch.`],
		['', 'Unload ("x")', '("x")', 'is a literal.'],
		['', 'Unload -1', '-1', 'is a literal.'],
		['s As String', 'Load s', 's', "Load takes an object, such as a UserForm, and 's' is declared As String."],
		['t As T1', 'Unload t', 't', "and 't' is declared As T1."],
		['o As Object', 'Put #f, , o', 'o', "Put cannot write an object, and 'o' is declared As Object. This is a VBE compile error: Can not Get or Put an object reference variable or a variable of user-defined type containing an object reference."],
		['c As New Collection', 'Get #f, 1, c', 'c', "Get cannot read into an object, and 'c' is declared As Collection."],
		['t3 As T3', 'Put #f, , t3', 't3', "and 't3' is declared As T3, a user-defined type holding an object."],
		['ta(2) As T2', 'Get #f, , ta', 'ta', "and 'ta' is an array of T2, a user-defined type holding an object."],
		['t3 As T3', 'Put #f, , t3.t.o', 't3.t.o', "and 't3.t.o' is declared As Object."],
		['', 'Put #f, , Nothing', 'Nothing', "and 'Nothing' is an object reference."],
		['t6 As T6', 'Put #f, , t6', 't6', "and 't6' is declared As T6, a user-defined type holding an object."],
	])('reports %s with %j', (dims, line, span, message) => {
		const src = procedure(dims, 'If False Then', '    ' + line, 'End If');
		expectDiagnostic(src, operandTypes(src), 'statement-operand-type', { span, message });
	});

	it('reads a one-line If, a colon, a line number and a DefType line', () => {
		for (const line of ['If f = 0 Then Line Input #f, n', 'If f = 0 Then f = 1 Else Put #f, , o', 'f = 1: Mid(n, 1) = "x"', '10 Unload n']) {
			expect(operandTypes(procedure('n As Long, o As Object', line)), line).toHaveLength(1);
		}
		const defLng = ['Option Explicit', 'DefLng X', 'Sub Main(ByRef p As Long)', '    Dim f As Integer, x', '    Line Input #f, x', '    Mid(x, 1) = "x"', '    Line Input #f, p', 'End Sub', ''].join('\r\n');
		expect(operandTypes(defLng)).toHaveLength(3);
	});

	it('leaves a type character that names another type to type-suffix-mismatch', () => {
		const src = procedure('n As Long', 'If False Then Line Input #f, n$');
		expect(operandTypes(src)).toEqual([]);
		expect(byCode(analyzeModule(src, { host: 'excel' }), 'type-suffix-mismatch')).toHaveLength(1);
	});

	it('leaves the operands that compile alone', () => {
		const lines = [
			'Line Input #f, s', 'Line Input #f, v', 'Line Input #f, x', 'Line Input #f, fs', 'Line Input #f, s$', 'Line Input #f, sa(1)', 'Line Input #f, t.s',
			'Mid(s, 1) = "x"', 'Mid$(v, 1, 1) = "x"', 'Mid(fs, 1) = n', 'Mid(t.s, 1) = o',
			'Unload v', 'Unload o', 'Unload c', 'Unload Nothing', 'Unload Empty', 'Unload Null', 'Unload oa(1)', 'Unload va(1)', 'Unload t2.o', 'Load o',
			'Put #f, , v', 'Get #f, , t', 'Put #f, , t4', 'Put #f, , t4.v', 'Get #f, , oa', 'Put #f, , oa()', 'Put #f, , va', 'Put #f, 1, s',
			'Mid(fsa(1), 1) = "x"', 'Unload (o)', 'Put #f, , t7.oa', 'Put #f, , t7.oa()',
		];
		const dims = 't7 As T7, fsa(2) As String * 3, s As String, v As Variant, x, fs As String * 5, sa(2) As String, n As Long, o As Object, c As New Collection, oa(2) As Object, va(2) As Variant, t As T1, t2 As T2, t4 As T4';
		for (const line of lines) {
			expect(operandTypes(procedure(dims, 'If False Then', '    ' + line, 'End If')), line).toEqual([]);
		}
	});

	it('reads each Type once, however often others hold it', () => {
		// T30 holds two T29s, each holding two T28s, down to T0: 2^30 paths, 31 Types.
		const types = ['Private Type T0', '    n As Long', 'End Type'];
		for (let i = 1; i <= 30; i++) {
			types.push(`Private Type T${i}`, `    a As T${i - 1}`, `    b As T${i - 1}`, 'End Type');
		}
		const source = ['Option Explicit', ...types, 'Private Type H', '    t As T30', '    o As Object', 'End Type', 'Sub Main()', '    Dim f As Integer, x As T30, h As H',
			'    Put #f, , x', '    Put #f, , h', 'End Sub', ''].join('\r\n');
		expect(operandTypes(source).map((d) => d.message)).toEqual([expect.stringContaining("'h' is declared As H, a user-defined type holding an object")]);
	});

	it('follows no field of a Type of another module, even one named as this module\'s', () => {
		const modules = [
			{ moduleName: 'Module1', type: 'standard', source: ['Option Explicit', 'Private Type T1', '    s As Long', 'End Type', 'Sub Main()', '    Dim f As Integer, x As Module2.T1',
				'    Line Input #f, x.s', 'End Sub', ''].join('\r\n') },
			{ moduleName: 'Module2', type: 'standard', source: 'Option Explicit\r\nPublic Type T1\r\n    s As String\r\nEnd Type\r\n' },
		];
		expect(projectErrors(modules, 'excel')).toEqual([]);
	});

	it('reads a Get or Put operand as an object only when it is known to be one', () => {
		// A library's Enum is no object, and nor is a Type holding one.
		const enums = ['Option Explicit', 'Private Type T8', '    a As VbMsgBoxResult', '    n As Long', 'End Type', 'Sub Main()',
			'    Dim f As Integer, r As VbMsgBoxResult, ra(2) As VbMsgBoxResult, t As T8, p As POINTAPI',
			'    Put #f, , r', '    Put #f, , ra(1)', '    Get #f, , t', '    Put #f, , p', 'End Sub', ''].join('\r\n');
		expect(operandTypes(enums)).toEqual([]);
		// A project class and a variable declared As New are.
		const classes = projectErrors([
			{ moduleName: 'Module1', type: 'standard', source: ['Option Explicit', 'Sub Main()', '    Dim f As Integer, k As Class1, d As New Class1', '    Put #f, , k', '    Get #f, , d', 'End Sub', ''].join('\r\n') },
			{ moduleName: 'Class1', type: 'class', source: 'Option Explicit\r\nPublic s As String\r\n' },
		], 'excel');
		expect(classes).toEqual(['statement-operand-type', 'statement-operand-type']);
		// A Collection is one, with New or not.
		expect(operandTypes(procedure('c As Collection, c2 As VBA.Collection', 'Put #f, , c', 'Put #f, , c2'))).toHaveLength(2);
		// As New makes an object of a class XLIDE does not know; the same type without New is unknown.
		const auto = ['Option Explicit', 'Sub Main()', '    Dim f As Integer, d As New Scripting.Dictionary, e As Scripting.Dictionary', '    Put #f, , d', '    Put #f, , e', 'End Sub', ''].join('\r\n');
		expect(operandTypes(auto).map((d) => d.message)).toEqual([expect.stringContaining("'d' is declared As Scripting.Dictionary.")]);
	});

	it('reads a variable of another module by its declared type', () => {
		const errors = projectErrors([
			{ moduleName: 'Module1', type: 'standard', source: ['Option Explicit', 'Sub Main()', '    Dim f As Integer', '    If False Then', '        Line Input #f, gN',
				'        Put #f, , gO', '        Unload gS', '    End If', 'End Sub', ''].join('\r\n') },
			{ moduleName: 'Module2', type: 'standard', source: 'Option Explicit\r\nPublic gN As Long, gO As Object, gS As String\r\n' },
		], 'excel');
		expect(errors).toEqual(['statement-operand-type', 'statement-operand-type', 'statement-operand-type']);
	});

	it('reads a variable of another module by its own module\'s Types', () => {
		const run = (main: string[], other: string[]) => projectErrors([
			{ moduleName: 'Module1', type: 'standard', source: ['Option Explicit', ...main, 'Sub Main()', '    Dim f As Integer', '    Line Input #f, gT.s', '    Put #f, , gT', '    Put #f, , h.inner', 'End Sub', ''].join('\r\n') },
			{ moduleName: 'Module2', type: 'standard', source: ['Option Explicit', ...other, 'Public gT As T1', ''].join('\r\n') },
		], 'excel');
		// Module1's T1 has a Long s and an object; Module2's, which gT is, a String s only.
		expect(run(['Private Type T1', '    s As Long', '    o As Object', 'End Type', 'Private Type H', '    inner As Module2.T1', 'End Type', 'Private h As H'],
			['Public Type T1', '    s As String', 'End Type'])).toEqual([]);
	});

	it('leaves another module\'s variable of a type named as one of this module\'s alone', () => {
		// gE is Module2's Enum T1, not Module1's Type T1.
		const errors = projectErrors([
			{ moduleName: 'Module1', type: 'standard', source: ['Option Explicit', 'Private Type T1', '    n As Long', 'End Type', 'Sub Main()', '    Dim f As Integer, s As String',
				'    If False Then', '        Mid(s, 1) = gE', '        Line Input #f, gE', '        Unload gE', '    End If', 'End Sub', ''].join('\r\n') },
			{ moduleName: 'Module2', type: 'standard', source: 'Option Explicit\r\nPublic Enum T1\r\n    E1A = 1\r\nEnd Enum\r\nPublic gE As T1\r\n' },
		], 'excel');
		expect(errors).toEqual([]);
	});

	it('reads a procedure\'s own name as its result', () => {
		const inner = (header: string, end: string, line: string) => operandTypes(['Option Explicit', header, '    Dim f As Integer', `    If False Then ${line}`, end, ''].join('\r\n'));
		expect(inner('Function Fn() As Long', 'End Function', 'Line Input #f, Fn')).toHaveLength(1);
		expect(inner('Function Fn() As Long', 'End Function', 'Unload Fn')).toHaveLength(1);
		expect(inner('Property Get P() As Object', 'End Property', 'Put #f, , P').map((d) => d.message)).toEqual([expect.stringContaining("'P' is declared As Object.")]);
		// With arguments the name is a call, not an element of the result.
		expect(inner('Function Fo(i As Long) As Object()', 'End Function', 'Put #f, , Fo(1)')).toEqual([]);
		expect(inner('Function Fa() As Long()', 'End Function', 'Line Input #f, Fa(1)')).toEqual([]);
		// Nor is a member of the result followed.
		const member = ['Option Explicit', 'Private Type T2', '    o As Object', '    n As Long', 'End Type', 'Function Ft() As T2', '    Dim f As Integer',
			'    If False Then Put #f, , Ft.o', '    If False Then Line Input #f, Ft.n', 'End Function', ''].join('\r\n');
		expect(operandTypes(member)).toEqual([]);
	});

	it('leaves a type qualified by another module and named as one of this module\'s Types alone', () => {
		const errors = projectErrors([
			{ moduleName: 'Module1', type: 'standard', source: ['Option Explicit', 'Private Type T1', '    n As Long', 'End Type', 'Private Type H', '    e As Module2.T1', 'End Type',
				'Sub Main()', '    Dim f As Integer, s As String, x As Module2.T1, h As H', '    If False Then', '        Mid(s, 1) = x', '        Mid(s, 1) = h.e', '        Put #f, , x', '    End If', 'End Sub', ''].join('\r\n') },
			{ moduleName: 'Module2', type: 'standard', source: 'Option Explicit\r\nPublic Type T1\r\n    o As Object\r\nEnd Type\r\n' },
		], 'excel');
		expect(errors.filter((code) => code === 'statement-operand-type' || code === 'udt-value-mismatch')).toEqual([]);
	});

	it('leaves a ParamArray under a DefType line to its declaration', () => {
		// The VBE refuses the declaration: ParamArray must be declared as an array of Variant.
		const src = ['Option Explicit', 'DefLng A-Z', 'Sub Main(ParamArray x())', '    Dim f As Integer', '    Line Input #f, x(0)', '    Mid(x(0), 1) = "x"', '    Unload x(0)', 'End Sub', ''].join('\r\n');
		expect(operandTypes(src)).toEqual([]);
	});

	it('stops at a Type that holds itself', () => {
		const src = ['Option Explicit', 'Private Type A', '    x As A', '    n As Long', 'End Type', 'Sub Main()', '    Dim f As Integer, t As A', '    Put #f, , t', 'End Sub', ''].join('\r\n');
		expect(operandTypes(src)).toEqual([]);
	});

	it('reads the statements in a class module, where its own Sub Unload is called', () => {
		const lines = (...body: string[]) => ['Option Explicit', 'Public Count As Long', 'Sub Run()', '    Dim f As Integer', ...body.map((line) => '    ' + line), 'End Sub', ''].join('\r\n');
		const inClass = (source: string) => byCode(analyzeModule(source, { host: 'excel', moduleKind: 'class' }), 'statement-operand-type');
		expect(inClass(lines('Line Input #f, Count', 'Unload "x"'))).toHaveLength(2);
		expect(inClass(lines('Unload Me') + 'Sub Unload(a)\r\nEnd Sub\r\n')).toEqual([]);
		expect(inClass(lines('Unload "x"') + 'Sub Unload(a)\r\nEnd Sub\r\n')).toEqual([]);
	});

	it('leaves a statement a name of the project takes the place of', () => {
		const shadowed = (main: string, other: string) => projectErrors([
			{ moduleName: 'Module1', type: 'standard', source: main },
			{ moduleName: 'Module2', type: 'standard', source: other },
		], 'excel');
		expect(shadowed(procedure('', 'Unload "x"', 'Load "x"'), 'Option Explicit\r\nSub Unload(a)\r\nEnd Sub\r\nSub Load(a)\r\nEnd Sub\r\n')).toEqual([]);
		// Another module's Private one is not seen.
		expect(shadowed(procedure('', 'If False Then Unload "x"'), 'Option Explicit\r\nPrivate Sub Unload(a)\r\nEnd Sub\r\n')).toEqual(['statement-operand-type']);
		// No name takes the Mid statement's place, in this module or another.
		expect(shadowed(procedure('n As Long', 'Mid(n, 1) = "x"'), 'Option Explicit\r\nFunction Mid(a, b)\r\nEnd Function\r\n')).toEqual(['statement-operand-type']);
		expect(operandTypes(procedure('n As Long', 'Mid(n, 1) = "x"') + 'Function Mid(a, b)\r\nEnd Function\r\n')).toHaveLength(1);
		expect(operandTypes(procedure('n As Long, Mid As Variant', 'Mid(n, 1) = "x"'))).toHaveLength(1);
	});
});
