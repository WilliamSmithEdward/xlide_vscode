// Diagnostics tests: objects and arrays read as a condition or a Boolean
// operand (issue #424). Measured in Excel 16.0 64-bit (2026-10-02) through
// pyVBAharness, each value in each form.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const FORMS = {
	oneLineIf: 'If x Then Main = 1',
	blockIf: 'If x Then\n        Main = 1\n    End If',
	elseIf: 'If False Then\n    ElseIf x Then\n        Main = 1\n    End If',
	doWhile: 'Do While x\n        Exit Do\n    Loop',
	loopUntil: 'Do\n        Main = 1\n    Loop Until x',
	iif: 'Main = IIf(x, 1, 2)',
	not: 'Main = Not x',
	select: 'Select Case x\n        Case 1\n            Main = 1\n    End Select',
	and: 'Main = x And True',
	while: 'While x\n        Main = 1\n    Wend',
};
type Form = keyof typeof FORMS;

function found(declaration: string, form: Form): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    ${declaration}\n    ${FORMS[form]}\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

const ALL = Object.keys(FORMS) as Form[];

describe('a value that cannot be a condition', () => {
	it('raises 91 on a never-set object in every form', () => {
		for (const form of ALL) {
			expect(found('Dim x As Object', form), form).toEqual(['object-variable-not-set']);
		}
	});

	it('raises 450 or does not compile on a Collection', () => {
		for (const form of ['oneLineIf', 'blockIf', 'elseIf', 'doWhile', 'loopUntil', 'iif', 'while'] as Form[]) {
			expect(found('Dim x As New Collection', form), form).toEqual(['object-default-value']);
		}
		for (const form of ['not', 'select', 'and'] as Form[]) {
			expect(found('Dim x As New Collection', form), form).toEqual(['collection-operand']);
		}
	});

	it('does not compile on an array, but for IIf (13) and Not, which runs', () => {
		for (const declaration of ['Dim x(1) As Long', 'Dim x() As Long']) {
			for (const form of ['oneLineIf', 'blockIf', 'elseIf', 'doWhile', 'loopUntil', 'select', 'and', 'while'] as Form[]) {
				expect(found(declaration, form), `${declaration} ${form}`).toEqual(['non-scalar-binary-operand']);
			}
			expect(found(declaration, 'iif')).toEqual(['variant-value-misuse']);
			expect(found(declaration, 'not')).toEqual([]);
		}
	});

	it('raises 13 on a Variant holding an array in every form', () => {
		for (const form of ALL) {
			expect(found('Dim x As Variant\n    x = Array(1)', form), form).toEqual(['variant-value-misuse']);
		}
	});
});

describe('what reads a value from them', () => {
	it('stays quiet', () => {
		const src = (body: string) => `Option Explicit\nFunction Main() As Variant\n    ${body}\nEnd Function\n`;
		for (const body of [
			'Dim x As Object\n    Set x = New Collection\n    If x Is Nothing Then Main = 1',
			'Dim x As New Collection\n    If x.Count Then Main = 1',
			'Dim x(1) As Long\n    If x(0) Then Main = 1',
			'Dim x As Variant\n    x = Array(1)\n    If x(0) Then Main = 1',
		]) {
			expect(analyzeModule(src(body)).filter((diag) => diag.severity === 'error'), body).toEqual([]);
		}
	});
});

describe('a String a loop condition cannot convert', () => {
	it('is read with what reaches the loop, though the body assigns it', () => {
		const src = 'Option Explicit\nFunction Main() As Variant\n    Dim x As String\n    x = "abc"\n    While x\n        x = "0"\n    Wend\n    Main = 1\nEnd Function\n';
		expect(analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code)).toEqual(['string-arithmetic-coercion']);
	});
});
