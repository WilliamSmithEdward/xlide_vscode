// Declarations the VBE refuses by their form (issue #212). Each verdict is a
// full project compile in 64-bit Excel 16.0: running a macro compiles only
// what it calls, and let several of these through.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode } from '../helpers/diagnostics';

const ENUM = 'Private Enum Mode\n    mA = 1\nEnd Enum\n';
const TYPE = 'Private Type T\n    x As Long\nEnd Type\n';
const codes = (src: string, moduleKind: 'standard' | 'class' = 'standard'): string[] =>
	analyzeModule(`Option Explicit\n${src}`, { moduleKind }).map((d) => d.code ?? '');

describe('private-type-in-public-signature (issue #212)', () => {
	const CODE = 'private-type-in-public-signature';
	it.each([
		['a Public Function parameter', `${ENUM}Public Function Pick(ByVal m As Mode) As Long\nEnd Function\n`],
		['an unmarked Sub, which is Public', `${ENUM}Sub S(v As Mode)\nEnd Sub\n`],
		['a Public Function result', `${ENUM}Public Function Pick() As Mode\nEnd Function\n`],
		['a Public Property Get result', `${ENUM}Public Property Get P() As Mode\nEnd Property\n`],
		['a Public variable', `${ENUM}Public v As Mode\n`],
		['a Global variable', `${ENUM}Global g As Mode\n`],
		['an array parameter', `${ENUM}Public Sub S(v() As Mode)\nEnd Sub\n`],
		['a Public Declare', `${ENUM}Public Declare PtrSafe Sub Sleep2 Lib "kernel32" Alias "Sleep" (ByVal ms As Mode)\n`],
	])('refuses a standard module s Private Enum in %s', (_name, src) => {
		expect(codes(src)).toContain(CODE);
	});

	it.each([
		['a Private Sub', `${ENUM}Private Sub S(v As Mode)\nEnd Sub\n`],
		['a module-level Dim', `${ENUM}Dim d As Mode\n`],
		['a Public Type field', `${ENUM}Public Type PT\n    m As Mode\nEnd Type\n`],
		['a Private Type in a Public Sub', `${TYPE}Public Sub S(v As T)\nEnd Sub\n`],
		['a Private Type as a Public Function result', `${TYPE}Public Function F() As T\nEnd Function\n`],
		['a Private Type as a Public variable', `${TYPE}Public v As T\n`],
		['a Private Type in a Public Sub beside a Private Enum', `${ENUM}${TYPE}Public Sub S(v As T)\nEnd Sub\n`],
	])('accepts in a standard module %s', (_name, src) => {
		expect(codes(src)).not.toContain(CODE);
	});

	it('refuses a class module s Private Enum or Type in a public procedure or Event', () => {
		expect(codes(`${TYPE}Public Sub S(v As T)\nEnd Sub\n`, 'class')).toContain(CODE);
		expect(codes(`${TYPE}Public Function F() As T\nEnd Function\n`, 'class')).toContain(CODE);
		expect(codes(`${ENUM}Public Sub S(v As Mode)\nEnd Sub\n`, 'class')).toContain(CODE);
		expect(codes(`${ENUM}Public Event E(ByVal m As Mode)\n`, 'class')).toContain(CODE);
		expect(codes(`${ENUM}Public v As Mode\n`, 'class')).toContain(CODE);
		expect(codes(`${TYPE}Friend Sub S(v As T)\nEnd Sub\n`, 'class')).not.toContain(CODE);
		expect(codes(`${TYPE}Private Sub S(v As T)\nEnd Sub\n`, 'class')).not.toContain(CODE);
	});

	it('leaves a class s Public variable of its Private Type, and a Global, to object-module-public-member', () => {
		expect(codes(`${TYPE}Public v As T\n`, 'class')).toEqual(['object-module-public-member']);
		expect(codes('Global g As Long\n', 'class')).toEqual(['object-module-public-member']);
	});
});

describe('optional-property-value and event-parameter-form (issue #212)', () => {
	it('refuses an Optional value on Property Let and Set, and not an Optional index', () => {
		expect(codes('Private Property Let Item(Optional ByVal v As Variant)\nEnd Property\n')).toContain('optional-property-value');
		expect(codes('Public Property Set Item(Optional ByVal v As Object)\nEnd Property\n')).toContain('optional-property-value');
		expect(codes('Public Property Let Item(Optional ByVal i As Long, ByVal v As Long)\nEnd Property\n')).not.toContain('optional-property-value');
	});

	it('refuses an Optional or ParamArray Event parameter', () => {
		expect(codes('Public Event Changed(Optional ByVal v As Long)\n', 'class')).toContain('event-parameter-form');
		expect(codes('Public Event Changed(ParamArray v())\n', 'class')).toContain('event-parameter-form');
		expect(codes('Public Event Changed(ByVal v As Long)\n', 'class')).not.toContain('event-parameter-form');
	});
});

describe('const-invalid-type (issue #212)', () => {
	const constMessage = (decl: string): string | undefined =>
		byCode(analyzeModule(decl), 'const-invalid-type')[0]?.message;

	it('refuses Object as an invalid type, and a class, Enum or Type as no type name', () => {
		expect(constMessage('Private Const C1 As Object = Nothing\n')).toContain('Invalid data type for constant');
		expect(constMessage('Sub T()\n    Const C1 As Object = Nothing\nEnd Sub\n')).toContain('Invalid data type for constant');
		for (const type of ['Collection', 'Range', 'Excel.Range', 'Class1', 'E']) {
			expect(constMessage(`Private Const C1 As ${type} = Nothing\n`), type).toContain('Expected: type name');
		}
	});

	it('accepts VBA s own types', () => {
		for (const decl of ['As Variant = Empty', 'As Date = #1/1/2000#', 'As Currency = 1', 'As LongPtr = 1', 'As String * 5 = "a"']) {
			expect(constMessage(`Private Const C1 ${decl}\n`), decl).toBeUndefined();
		}
	});
});

describe('empty-enum, type-member-without-type and type-enum-name-conflict (issue #212)', () => {
	it('refuses an Enum with no members', () => {
		expect(codes('Private Enum E\nEnd Enum\n')).toContain('empty-enum');
		expect(codes('Public Enum E\n    a\nEnd Enum\n')).not.toContain('empty-enum');
	});

	it('refuses a Type member with no As clause, a type suffix or a line number included', () => {
		for (const field of ['a', 'a(1)', 'a&']) {
			expect(codes(`Private Type T\n    ${field}\nEnd Type\n`), field).toContain('type-member-without-type');
		}
		const numbered = byCode(analyzeModule('Private Type Customer\n10  Id As Long\nEnd Type\n'), 'type-member-without-type');
		expect(numbered[0].message).toContain('A line number cannot label a Type member');
	});

	it('refuses a Type and an Enum of one name, and not a Const sharing either', () => {
		expect(codes(`${TYPE}Private Enum T\n    tA = 1\nEnd Enum\n`)).toContain('type-enum-name-conflict');
		expect(codes(`${TYPE}Private Const T As Long = 1\n`)).not.toContain('type-enum-name-conflict');
	});
});

describe('redim-type-change (issue #212)', () => {
	const redim = (decls: string, body: string) => codes(`${decls}Sub T2()\n${body}End Sub\n`).filter((code) => code === 'redim-type-change');

	it('refuses a ReDim that gives an array another element type', () => {
		expect(redim('', '    Dim x() As String\n    ReDim x(1) As Long\n')).toHaveLength(1);
		expect(redim('Private m() As String\n', '    ReDim m(1) As Long\n')).toHaveLength(1);
		expect(redim('', '    Dim v()\n    ReDim v(1) As Long\n')).toHaveLength(1);
		expect(redim('', '    Dim x() As String\n    ReDim x(1)\n    ReDim Preserve x(2) As Long\n')).toHaveLength(1);
	});

	it('accepts the same type, and any type for a Variant that is not an array', () => {
		expect(redim('', '    Dim x() As String\n    ReDim x(1) As String\n')).toHaveLength(0);
		expect(redim('', '    Dim v As Variant\n    ReDim v(1) As Long\n')).toHaveLength(0);
	});
});

it('does not universally label an Optional setter value Syntax error',()=>{
 const source='Property Let State(Optional ByVal value As Boolean = False)\nEnd Property\nSub T()\nState = True\nEnd Sub';
 const finding=analyzeModule(source).find(d=>d.code==='optional-property-value');
 expect(finding?.message).toContain('VBE compile error');
 expect(finding?.message).not.toContain('Syntax error');
});
