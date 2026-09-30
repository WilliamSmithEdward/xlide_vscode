// Statements the VBE refuses by their form or by the type of what they name
// (issue #213). Each verdict is a full project compile in 64-bit Excel 16.0.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const TYPES = 'Private Type T\n    a As Long\nEnd Type\nPrivate Type U\n    a As Long\nEnd Type\n';
const inSub = (body: string, decls = '', moduleKind: 'standard' | 'class' = 'standard'): string[] =>
	analyzeModule(`Option Explicit\n${decls}Sub P1()\n${body}End Sub\n`, { moduleKind }).map((d) => d.code ?? '');

describe('for-variable-in-use and for-counter-type (issue #213)', () => {
	it('refuses a For or For Each nested on the same control variable', () => {
		expect(inSub('    Dim i As Long\n    For i = 1 To 2\n        For i = 1 To 2\n        Next\n    Next\n')).toContain('for-variable-in-use');
		expect(inSub('    Dim v As Variant\n    For Each v In Array(1)\n        For Each v In Array(2)\n        Next\n    Next\n')).toContain('for-variable-in-use');
		expect(inSub('    Dim v As Variant\n    For v = 1 To 2\n        For Each v In Array(2)\n        Next\n    Next\n')).toContain('for-variable-in-use');
	});

	it('accepts the same variable in loops one after the other', () => {
		expect(inSub('    Dim i As Long\n    For i = 1 To 2\n    Next\n    For i = 1 To 2\n    Next\n')).not.toContain('for-variable-in-use');
	});

	it.each(['String', 'Object', 'Boolean', 'Collection', 'T'])('refuses a For counter As %s', (type) => {
		expect(inSub(`    Dim c As ${type}\n    For c = 0 To 1\n    Next\n`, TYPES)).toContain('for-counter-type');
	});

	it.each(['Variant', 'Date', 'Long', 'Double', 'Currency'])('accepts a For counter As %s', (type) => {
		expect(inSub(`    Dim c As ${type}\n    For c = 0 To 1\n    Next\n`)).not.toContain('for-counter-type');
	});
});

describe('for-each-source-type on arrays of user-defined types (issue #213)', () => {
	it.each([
		['a fixed UDT array', '    Dim a(1) As T\n'],
		['a dynamic UDT array', '    Dim a() As T\n    ReDim a(1)\n'],
		['a fixed-length string array', '    Dim a(1) As String * 5\n'],
	])('refuses For Each over %s', (_name, decl) => {
		expect(inSub(`${decl}    Dim v As Variant\n    For Each v In a\n    Next\n`, TYPES)).toContain('for-each-source-type');
	});
});

describe('raiseevent-argument-count (issue #213)', () => {
	const raise = (event: string, call: string) => inSub(`    ${call}\n`, `${event}\n`, 'class').filter((code) => code === 'raiseevent-argument-count');

	it('refuses too many and too few arguments', () => {
		expect(raise('Public Event Changed(ByVal v As Long)', 'RaiseEvent Changed(1, 2)')).toHaveLength(1);
		expect(raise('Public Event Ev(ByVal a As Long)', 'RaiseEvent Ev')).toHaveLength(1);
	});

	it('accepts the right number', () => {
		expect(raise('Public Event Ev(ByVal a As Long)', 'RaiseEvent Ev(1)')).toHaveLength(0);
		expect(raise('Public Event Ev()', 'RaiseEvent Ev')).toHaveLength(0);
		expect(raise('Public Event Ev(ByVal a As Long, ByVal b As Long)', 'RaiseEvent Ev(Array(1, 2), 3)')).toHaveLength(0);
	});
});

describe('statement-before-first-case and return-with-value (issue #213)', () => {
	it('refuses a statement or label before the first Case, and not a comment', () => {
		expect(inSub('    Dim r As Long\n    Select Case 1\n        r = 2\n    Case 1\n    End Select\n')).toContain('statement-before-first-case');
		expect(inSub('    Select Case 1\nL1:\n    Case 1\n    End Select\n')).toContain('statement-before-first-case');
		expect(inSub("    Select Case 1\n        ' note\n    Case 1\n    End Select\n")).not.toContain('statement-before-first-case');
	});

	it('refuses Return with a value, and not a bare Return', () => {
		expect(analyzeModule('Function F() As Variant\n    Return 5\nEnd Function\n').map((d) => d.code)).toContain('return-with-value');
		expect(analyzeModule('Function F() As Variant\n    F = 1\n    Exit Function\n    Return\nEnd Function\n').map((d) => d.code)).not.toContain('return-with-value');
	});
});

describe('assignments the VBE refuses (issue #213)', () => {
	it('refuses assigning an Enum member', () => {
		expect(inSub('    eA = 2\n', 'Private Enum E\n    eA = 1\nEnd Enum\n')).toContain('const-assignment');
	});

	it('refuses assigning a Sub, or a Function of a type of VBA s own, from outside it', () => {
		expect(inSub('    Other = 2\n', 'Private Sub Other()\nEnd Sub\n')).toContain('assignment-to-procedure-name');
		expect(inSub('    F = 2\n', 'Private Function F() As Long\nEnd Function\n')).toContain('assignment-to-procedure-name');
		expect(inSub('    F = "x"\n', 'Private Function F() As String\nEnd Function\n')).toContain('assignment-to-procedure-name');
		expect(inSub('    F = 2\n', 'Private Function F() As Variant\nEnd Function\n')).not.toContain('assignment-to-procedure-name');
		expect(inSub('    Set F = Nothing\n', 'Private Function F() As Object\nEnd Function\n')).not.toContain('assignment-to-procedure-name');
		// Inside the Function its name is the return value.
		expect(analyzeModule('Function F() As Long\n    F = 2\nEnd Function\n').map((d) => d.code)).not.toContain('assignment-to-procedure-name');
	});

	it('refuses Set of a user-defined type', () => {
		expect(inSub('    Dim t As T\n    Set t = Nothing\n', TYPES)).toContain('set-requires-object');
	});

	it('refuses a number as a Mid target, and not a variable', () => {
		expect(inSub('    Mid(5, 1) = "x"\n')).toContain('mid-statement-literal-target');
		expect(inSub('    Dim s As String\n    s = "abc"\n    Mid(s, 1) = "x"\n')).not.toContain('mid-statement-literal-target');
	});
});

describe('with-scalar-target and paramarray-passing-mode (issue #213)', () => {
	it('refuses With on a scalar or a literal', () => {
		for (const setup of ['    Dim x As Long\n    With x\n', '    With "abc"\n', '    Dim s As String\n    With s\n', '    With 5\n']) {
			expect(inSub(`${setup}    End With\n`), setup).toContain('with-scalar-target');
		}
	});

	it('accepts With on a user-defined type or a Variant', () => {
		expect(inSub('    Dim t As T\n    With t\n    End With\n', TYPES)).not.toContain('with-scalar-target');
		expect(inSub('    Dim v As Variant\n    With v\n    End With\n')).not.toContain('with-scalar-target');
	});

	it('refuses ByVal or ByRef on a ParamArray', () => {
		expect(analyzeModule('Private Sub PB(ByVal ParamArray a())\nEnd Sub\n').map((d) => d.code)).toContain('paramarray-passing-mode');
		expect(analyzeModule('Private Sub PB(ByRef ParamArray a())\nEnd Sub\n').map((d) => d.code)).toContain('paramarray-passing-mode');
		expect(analyzeModule('Private Sub PB(ParamArray a())\nEnd Sub\n').map((d) => d.code)).not.toContain('paramarray-passing-mode');
	});
});

describe('udt-value-mismatch and udt-variant-coercion (issue #213)', () => {
	const udt = (body: string) => inSub(body, TYPES).filter((code) => code.startsWith('udt-'));

	it.each([
		['a number into a user-defined type', '    Dim t As T\n    t = 1\n'],
		['a comparison of two', '    Dim t As T, u As T\n    Debug.Print t = u\n'],
		['one printed', '    Dim t As T\n    Debug.Print t\n'],
		['one of another Type', '    Dim a As T, b As U\n    a = b\n'],
		['one into a Long', '    Dim n As Long, t As T\n    n = t\n'],
		['one in arithmetic', '    Dim t As T\n    Debug.Print t + 1\n'],
		['one as an If condition', '    Dim t As T\n    If t Then\n    End If\n'],
	])('refuses %s as a Type mismatch', (_name, body) => {
		expect(udt(body)).toContain('udt-value-mismatch');
	});

	it('refuses handing one to a Variant', () => {
		expect(udt('    Dim v As Variant, t As T\n    v = t\n')).toEqual(['udt-variant-coercion']);
		expect(udt('    Dim t As T\n    MsgBox t\n')).toEqual(['udt-variant-coercion']);
	});

	it('accepts one of the same Type, a member, an array element, and LSet', () => {
		expect(udt('    Dim a As T, b As T\n    a = b\n')).toEqual([]);
		expect(udt('    Dim t As T\n    t.a = 1\n')).toEqual([]);
		expect(udt('    Dim arr(1) As T, t As T\n    arr(0) = t\n    t = arr(1)\n')).toEqual([]);
		expect(udt('    Dim a As T, b As U\n    LSet a = b\n')).toEqual([]);
	});
});
