// Diagnostics tests: the members of a user-defined type (issue #253). Every
// case was measured in Excel 16.0 (build 20326, 2026-10-01).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

const TYPES = [
	'Private Type T',
	'    a As Long',
	'    i As Integer',
	'    dt As Date',
	'    o As Collection',
	'    d() As Long',
	'    f(0 To 2) As Long',
	'    v(1 To 2) As Integer',
	'    s As String',
	'End Type',
	'Private Type U',
	'    a As Long',
	'End Type',
].join('\n');

const PROCEDURES = [
	'Private Sub TakeV(ByVal v As Variant)',
	'End Sub',
	'Private Sub TakeOpt(Optional v As Variant)',
	'End Sub',
	'Private Sub TakeArr(ParamArray v() As Variant)',
	'End Sub',
	'Private Sub TakeT(x As T)',
	'End Sub',
	'Private Sub Fill(ByRef p As T)',
	'    Set p.o = New Collection',
	'    p.a = 2',
	'End Sub',
	'Private Sub FillO(ByRef o As Collection)',
	'    Set o = New Collection',
	'End Sub',
	'Private Sub Bump(ByRef n As Long)',
	'    n = 2',
	'End Sub',
	'Private Sub ResetI(ByRef p As T)',
	'    p.i = 0',
	'End Sub',
].join('\n');

function source(lines: readonly string[]): string {
	return `Option Explicit\n${TYPES}\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n${PROCEDURES}\n`;
}

function errors(lines: readonly string[]): string[] {
	return analyzeModule(source(lines)).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

function expectReport(lines: readonly string[], code: string, span: string, message: string | readonly string[]): void {
	const src = source(lines);
	expectDiagnostic(src, byCode(analyzeModule(src), code), code, { span, message });
}

describe('an object member nothing has Set', () => {
	it.each([
		[['Dim t As T', 't.o.Add 1'], 't.o'],
		[['Dim t As T', 'Main = t.o.Count'], 't.o'],
		[['Dim t As T', 'Main = t.o(1)'], 't.o'],
		[['Dim t As T', 'With t', '    .o.Add 1', 'End With'], '.o'],
		[['Dim t As T', 'Set t.o = New Collection', 'Set t.o = Nothing', 't.o.Add 1'], 't.o'],
	])('raises 91: %j', (lines, span) => {
		expectReport(lines, 'object-variable-not-set', span, ["Run-time error '91'", 'nothing has Set']);
	});

	it('raises 91 through With t.o', () => {
		expectReport(['Dim t As T', 'With t.o', '    .Add 1', 'End With'], 'object-variable-not-set', '.Add', "The With object 't.o'");
	});

	it('stays quiet once it is Set, or anything may have Set it', () => {
		for (const lines of [
			['Dim t As T', 'Set t.o = New Collection', 't.o.Add 1', 'Main = t.o.Count'],
			['Dim t As T', 'Main = (t.o Is Nothing)'],
			['Static t As T', 'Main = 1', 'If Not t.o Is Nothing Then Main = t.o.Count'],
			['Dim t As T', 'With t', '    Set .o = New Collection', 'End With', 't.o.Add 1'],
			['Dim t As T', 'Fill t', 'Main = t.o.Count'],
			['Dim t As T', 'FillO t.o', 'Main = t.o.Count'],
			['Dim t As T', 'Main = 0', 'If Not t.o Is Nothing Then', '    Main = t.o.Count', 'End If'],
			['Dim t As T', 'Main = 0', 'If Not t.o Is Nothing Then Main = t.o.Count'],
			['Dim t As T', 'Dim k As Long', 'For k = 1 To 2', '    If k = 2 Then Main = t.o.Count', '    Set t.o = New Collection', 'Next k'],
			// A label may be reached with the member Set: this runs.
			['Dim t As T', 'GoTo Skip', 'Again:', 't.o.Add 1', 'Main = t.o.Count', 'Exit Function', 'Skip:', 'Set t.o = New Collection', 'GoTo Again'],
		]) {
			expect(errors(lines), lines.join(': ')).not.toContain('object-variable-not-set');
		}
	});
});

describe('a numeric member', () => {
	it.each([
		[['Dim t As T', 'Main = 1 / t.a'], '/'],
		[['Dim t As T', 'Main = 1 \\ t.i'], '\\'],
		[['Dim t As T', 'Main = 5 Mod t.a'], 'Mod'],
	])('is 0 until assigned: %j', (lines, operator) => {
		expectReport(lines, 'division-by-zero', 't.' + (operator === '\\' ? 'i' : 'a'), `'${operator}' with a zero divisor`);
	});

	it('holds the literal it was last given', () => {
		expectReport(['Dim t As T', 't.a = 5', 't.a = 0', 'Main = 1 / t.a'], 'division-by-zero', 't.a', 'zero divisor');
	});

	it('stays quiet once assigned, or passed where it may be', () => {
		for (const lines of [
			['Dim t As T', 't.a = 2', 'Main = 1 / t.a'],
			['Dim t As T', 'With t', '    .a = 2', 'End With', 'Main = 1 / t.a'],
			['Dim t As T', 'Fill t', 'Main = 1 / t.a'],
			['Dim t As T', 'Bump t.a', 'Main = 1 / t.a'],
			['Dim t As T', 'Dim n As Long', 'n = 2', 't.a = n', 'Main = 1 / t.a'],
			['Dim t As T', 'Main = 0', 'If t.a <> 0 Then Main = 1 / t.a'],
		]) {
			expect(errors(lines), lines.join(': ')).not.toContain('division-by-zero');
		}
	});

	it.each([
		[['Dim t As T', 't.i = 32767', 't.i = t.i + 1'], 't.i + 1', '32768, outside the Integer range'],
		[['Dim t As T', 't.a = 40000', 't.i = t.a'], 't.a', "Assignment to 't.i' stores 40000 in an Integer"],
		[['Dim t As T', 't.dt = #12/31/9999#', 't.dt = t.dt + 1'], 't.dt + 1', 'outside the Date range'],
		[['Dim v(1 To 2) As Integer', 'v(2) = 40000'], '40000', "an element of 'v' stores 40000"],
		[['Dim t As T', 't.v(2) = 40000'], '40000', "an element of 't.v'"],
		[['Dim v() As Integer', 'ReDim v(2)', 'v(2) = 40000'], '40000', "an element of 'v'"],
	])('overflows: %j', (lines, span, message) => {
		expectReport(lines, 'arithmetic-overflow', span, message);
	});

	it('stays quiet in range, in a Variant element, and once the Type is passed whole', () => {
		expect(errors(['Dim t As T', 't.i = 32766', 't.i = t.i + 1', 'Main = t.i'])).toEqual([]);
		expect(errors(['Dim t As T', 't.i = 32767', 'ResetI t', 't.i = t.i + 1', 'Main = t.i'])).toEqual([]);
		expect(errors(['Dim v(1 To 2) As Integer', 'v(2) = 32767', 'Main = v(2)'])).toEqual([]);
		expect(errors(['Dim v(2) As Variant', 'v(2) = 40000', 'Main = v(2)'])).toEqual([]);
	});
});

describe('what the VBE refuses on a member or a whole Type', () => {
	it.each([
		[['Dim t As T', 'Main = t.a.Value'], 'scalar-member-access', 't.a.', "'t.a' is invalid because it is declared as Long"],
		[['Dim t As T', 'Main = t.s.Length'], 'scalar-member-access', 't.s.', 'declared as String'],
		[['Dim t As T', 'With t', '    Main = .a.Value', 'End With'], 'scalar-member-access', '.a.', 'Invalid qualifier'],
		[['Dim t As T', 'Erase t'], 'erase-requires-array', 't', 'Expected array'],
		[['Dim t As T', 'Erase t.a'], 'erase-requires-array', 't.a', 'declared As Long'],
		[['Dim t As T', 'ReDim t.f(3)'], 'fixed-array-redim', 't.f', 'Array already dimensioned'],
	])('%j', (lines, code, span, message) => {
		expectReport(lines, code, span, message);
	});

	it('stays quiet on what compiles', () => {
		expect(errors(['Dim t As T', 'ReDim t.d(3)', 'Main = 1'])).toEqual([]);
		expect(errors(['Dim ts(1) As T', 'Erase ts', 'Main = 1'])).toEqual([]);
	});
});

describe('a Type passed where VBA needs a Variant', () => {
	it.each([
		['TakeV t'], ['TakeOpt t'], ['TakeArr t'], ['Call TakeV(t)'],
		['Main = IsEmpty(t)'], ['Main = TypeName(t)'], ['Main = VarType(t)'], ['Main = IsNull(t)'],
		['Main = CVar(t)'], ['Main = Format(t)'], ['Main = IIf(True, 1, t)'], ['Main = Choose(1, 1, t)'],
		['Main = VBA.TypeName(t)'],
	])('is refused: %s', (line) => {
		expectReport(['Dim t As T', line], 'udt-variant-coercion', 't', 'Only user-defined types defined in public object modules');
	});

	it.each([
		[['Dim t As T', 'Dim c As New Collection', 'c.Add t']],
		[['Dim t As T', 'Dim c As New Collection', 'c.Add Item:=t']],
		[['Dim t As T', 'Dim v As Variant', 'v = Array(t)']],
		[['Dim t As T', 'Dim o As Object', 'Set o = New Collection', 'o.Add t']],
	])('is refused by a Collection or a late-bound call: %j', (lines) => {
		expect(errors(lines)).toEqual(['udt-variant-coercion']);
	});

	it('is a Type mismatch for CStr', () => {
		expectReport(['Dim t As T', 'Main = CStr(t)'], 'udt-value-mismatch', 't', 'Type mismatch');
	});

	it('stays quiet on Len, LenB, VarPtr, a member, and a Type parameter', () => {
		for (const line of ['Main = Len(t)', 'Main = LenB(t)', 'Main = (VarPtr(t) <> 0)', 'TakeV t.a', 'TakeT t', 'Call TakeT(t)']) {
			expect(errors(['Dim t As T', line, 'Main = 1']), line).toEqual([]);
		}
	});
});

describe('a Type parameter', () => {
	it.each([
		[['Dim v As Variant', 'TakeT v'], 'v', "'v' is declared As Variant"],
		[['Dim v', 'TakeT v'], 'v', 'Variant'],
		[['Dim u As U', 'TakeT u'], 'u', "'u' is declared As U"],
		[['Dim n As Long', 'TakeT n'], 'n', 'Long'],
		[['Dim c As Collection', 'TakeT c'], 'c', 'Collection'],
		[['TakeT 5'], '5', '5 is a literal'],
		[['Dim t As T', 'TakeT t.a'], 't.a', "'t.a' is a member declared As Long"],
	])('takes only a variable of its Type: %j', (lines, span, message) => {
		expectReport(lines, 'byref-argument-type-mismatch', span, ['ByRef argument type mismatch', message]);
	});

	it.each([['TakeT (t)'], ['Call TakeT((t))']])('takes no Type in parentheses: %s', (line) => {
		expectReport(['Dim t As T', line], 'variable-required', 't', 'Variable required');
	});
});

describe('LSet between two Types', () => {
	const lset = (dst: string, src: string): string[] => {
		const types = `Private Type A\n    ${dst}\n    y As Long\nEnd Type\nPrivate Type B\n    ${src}\n    z As Long\nEnd Type`;
		return analyzeModule(`Option Explicit\n${types}\nFunction Main() As Variant\n    Dim dst As A, src As B\n    LSet dst = src\n    Main = 1\nEnd Function\n`)
			.filter((diag) => diag.severity === 'error').map((diag) => diag.code);
	};

	it.each(['x As String', 'x() As Long', 'x As Collection', 'x As Variant'])('is refused when either holds %s', (member) => {
		expect(lset(member, 'x As Long')).toEqual(['lset-type-mismatch']);
		expect(lset('x As Long', member)).toEqual(['lset-type-mismatch']);
		expect(lset(member, member)).toEqual(['lset-type-mismatch']);
	});

	it.each(['x As Long', 'x As String * 4', 'x(1) As Long'])('takes two Types of fixed-size members: %s', (member) => {
		expect(lset(member, member)).toEqual([]);
		expect(lset(member, 'x As Long')).toEqual([]);
	});

	it('takes two values of one Type', () => {
		const src = 'Option Explicit\nPrivate Type A\n    x As String\n    o As Collection\nEnd Type\nFunction Main() As Variant\n    Dim p As A, q As A\n    LSet p = q\n    Main = 1\nEnd Function\n';
		expect(analyzeModule(src).filter((diag) => diag.severity === 'error')).toEqual([]);
	});
});

describe('a Declare taking a Type', () => {
	it('refuses it ByVal and Optional', () => {
		const declare = (param: string): string[] => analyzeModule(`Option Explicit\nPrivate Type T\n    a As Long\nEnd Type\nPrivate Declare PtrSafe Sub S Lib "kernel32" Alias "Sleep" (${param})\n`)
			.map((diag) => diag.code);
		expect(declare('ByVal x As T')).toContain('byval-udt-parameter');
		expect(declare('Optional x As T')).toContain('optional-udt-parameter');
		expect(declare('ByRef x As T')).not.toContain('byval-udt-parameter');
	});
});
