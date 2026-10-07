// Diagnostics tests: a project class's member misused through a variable
// declared As the class (issue #414). Measured in Excel 16.0 64-bit
// (2026-10-02) through pyVBAharness; each is a compile error.

import { describe, it, expect } from 'vitest';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';
import { buildVbaProjectIndex, projectAnalysisOptionsForModule, projectProcedureSignatures } from '../../src/vbaProjectAnalysis';

function found(member: string, body: string): string[] {
	const modules = [
		{ moduleName: 'Module1', type: 'standard', source: `Option Explicit\nFunction Main() As Variant\n    Dim c As New Class1\n    ${body}\nEnd Function\n` },
		{ moduleName: 'Class1', type: 'class', source: `Option Explicit\n${member}\n` },
	];
	const project = buildVbaProjectIndex(modules, undefined, { conditionalCompilation: { projectConstants: {} } });
	const procedures = projectProcedureSignatures(project);
	const { diagnostics } = analyzeVbaModuleSource({
		source: modules[0].source,
		moduleName: 'Module1',
		moduleKind: 'standard',
		...projectAnalysisOptionsForModule(project, 'Module1', procedures),
	} as never);
	return diagnostics.filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

const SUB = 'Public Sub M()\nEnd Sub';
const LONG_GET = 'Public Property Get M() As Long\n    M = 1\nEnd Property';
const STR_GET = 'Public Property Get M() As String\n    M = "a"\nEnd Property';
const FN_ARG = 'Public Function M(ByVal i As Long) As Variant\n    M = i\nEnd Function';

describe('a class member used in a form the VBE refuses', () => {
	it('is Expected Function or variable on a Sub assigned or read by a Set', () => {
		for (const body of ['c.M = 5', 'Set c.M = New Collection', 'Dim o As Object\n    Set o = c.M', 'Main = c.M.Count', 'With c\n        .M = 1\n    End With']) {
			expect(found(SUB, body), body).toEqual(['sub-used-as-value']);
		}
	});

	it('is Invalid qualifier or Type mismatch on a scalar', () => {
		expect(found('Public M As Long', 'c.M.Add 1')).toEqual(['scalar-member-access']);
		expect(found(LONG_GET, 'Main = c.M.Count')).toEqual(['scalar-member-access']);
		expect(found(LONG_GET, 'Main = c.M Is Nothing')).toEqual(['is-operator-non-object']);
	});

	it('refuses an argument list a Get does not take, and a write to a String Get', () => {
		expect(found(LONG_GET, 'Main = c.M(1)')).toEqual(['argument-count']);
		expect(found(STR_GET, 'c.M(1) = 2')).toEqual(['readonly-member-assignment']);
	});

	it('is Argument not optional on a Function read without its argument', () => {
		expect(found(FN_ARG, 'Main = c.M')).toEqual(['argument-count']);
		expect(found(FN_ARG, 'Main = c.M & "x"')).toEqual(['argument-count']);
	});

	it('is Invalid use of property on a member of a Let-only property', () => {
		expect(found('Public Property Let M(ByVal v As Variant)\nEnd Property', 'c.M.Add 1')).toEqual(['invalid-property-use']);
	});

	it('is Argument not optional on a Collection field beside an operator', () => {
		expect(found('Public M As Collection', 'Main = c.M & "x"')).toEqual(['collection-operand']);
	});
});

describe('a class member used as its declaration allows', () => {
	it('stays quiet', () => {
		expect(found('Public M As Long', 'c.M = 5\n    Main = c.M')).toEqual([]);
		expect(found(LONG_GET, 'Main = c.M + 1')).toEqual([]);
		expect(found(FN_ARG, 'Main = c.M(2)')).toEqual([]);
		expect(found('Public M As Collection', 'Set c.M = New Collection\n    c.M.Add 1\n    Main = c.M.Count')).toEqual([]);
		expect(found(SUB, 'c.M\n    Main = 1')).toEqual([]);
	});

	it('takes a chain or a With argument in a call statement as a call', () => {
		const members = `Public Function Self() As Class1\n    Set Self = Me\nEnd Function\n${SUB.replace('M()', 'S(ByVal v As Variant)')}\n${FN_ARG.replace('M(', 'F(').replace('M = i', 'F = i')}`;
		expect(found(members, 'c.Self.F 1\n    c.Self.S 1\n    With c\n        .Self.F 2\n        c.S .Self\n        c.F .Self\n    End With\n    Main = 1')).toEqual([]);
		expect(found(members, 'Main = c.Self.F')).toEqual(['argument-count']);
		// A call in a single-line If's branch, and a chain with an indexed step.
		const indexed = `${members}\nPublic Function At(ByVal i As Long) As Class1\n    Set At = Me\nEnd Function`;
		expect(found(indexed, 'If Main Then c.F 1')).toEqual([]);
		expect(found(indexed, 'c.At(1).F 2')).toEqual([]);
		expect(found(indexed, 'c.Self.At(1).F 2')).toEqual([]);
		expect(found(members, 'Main = c.Self.F.Count')).toEqual(['argument-count']);
	});
});

describe('a Let into a class Function or a Set-only property (issue #414)', () => {
	const VARIANT_FN = 'Public Function M() As Variant\n    M = 1\nEnd Function';
	it('reports a Let into a Function returning a Variant or a Collection', () => {
		expect(found(VARIANT_FN, 'c.M = 5')).toEqual(['variant-value-misuse']);
		expect(found(VARIANT_FN, 'With c\n        .M = 1\n    End With')).toEqual(['variant-value-misuse']);
		expect(found('Public Function M()\n    M = 1\nEnd Function', 'c.M = 5')).toEqual(['variant-value-misuse']);
		expect(found('Public Function M() As Collection\n    Set M = New Collection\nEnd Function', 'c.M = 5')).toEqual(['argument-count']);
		expect(found(VARIANT_FN, 'c.M(1) = 2')).toEqual(['variant-value-misuse']);
	});

	it('stays quiet for object results and checks a known scalar despite optional parameters', () => {
		expect(found('Public Function M() As Variant\n    Set M = New Collection\nEnd Function', 'c.M = 5')).toEqual([]);
		expect(found('Public Function M(Optional ByVal i As Long) As Variant\n    M = 1\nEnd Function', 'c.M = 5')).toEqual(['variant-value-misuse']);
	});

	it('reports a Let into a Property Set with no Let, indexed or not', () => {
		const SET_ONLY = 'Public Property Set M(ByVal v As Object)\nEnd Property';
		expect(found(SET_ONLY, 'c.M(1) = 2')).toEqual(['invalid-property-use']);
		expect(found(SET_ONLY, 'c.M = 5')).toEqual(['set-required']);
		expect(found(SET_ONLY, 'Set c.M = New Collection')).toEqual([]);
	});

	it('stays quiet on a read, a Let property and a Function returning Object', () => {
		expect(found(VARIANT_FN, 'Main = c.M')).toEqual([]);
		expect(found('Public Property Let M(ByVal v As Variant)\nEnd Property', 'c.M = 5')).toEqual([]);
		expect(found('Public Function M() As Object\n    Set M = Nothing\nEnd Function', 'c.M = 5')).toEqual([]);
	});
});
