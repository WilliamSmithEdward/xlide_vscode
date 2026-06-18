// Diagnostics tests: argument-shape-mismatch rule (XLIDE v2.5.0).
// Mirrors the promoted argshape_* VBE oracle cases: an array / same-module Type
// argument into a scalar parameter, or a scalar into an array parameter, is a
// compile error; the matching / Variant / ParamArray controls stay quiet.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

import { byCode, expectDiagnostic } from '../helpers/diagnostics';

const CODE = 'argument-shape-mismatch';

describe('analyzeModule - argument-shape-mismatch', () => {
	it('flags a whole array passed to a ByRef scalar parameter', () => {
		const src =
			'Public Sub Mutate(ByRef value As Long)\n' +
			'End Sub\n' +
			'Public Sub Caller()\n' +
			'    Dim arr(0 To 2) As Long\n' +
			'    Mutate arr\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), CODE, {
			severity: 'error',
			span: 'arr',
			message: 'ByRef argument type mismatch',
		});
	});

	it('flags a whole array passed to a ByVal scalar parameter', () => {
		const src =
			'Public Sub Mutate(ByVal value As Long)\n' +
			'End Sub\n' +
			'Public Sub Caller()\n' +
			'    Dim arr(0 To 2) As Long\n' +
			'    Mutate arr\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), CODE, { span: 'arr', message: 'Type mismatch' });
	});

	it('flags a user-defined Type value passed to a scalar parameter', () => {
		const src =
			'Private Type T\n' +
			'    a As Long\n' +
			'End Type\n' +
			'Public Sub Mutate(ByRef value As Long)\n' +
			'End Sub\n' +
			'Public Sub Caller()\n' +
			'    Dim u As T\n' +
			'    Mutate u\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), CODE, { span: 'u', message: 'user-defined Type' });
	});

	it('flags a scalar passed to an array parameter', () => {
		const src =
			'Public Sub Mutate(ByRef value() As Long)\n' +
			'End Sub\n' +
			'Public Sub Caller()\n' +
			'    Dim x As Long\n' +
			'    Mutate x\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), CODE, {
			span: 'x',
			message: 'array or user-defined type expected',
		});
	});

	it('flags a Variant scalar passed to an array parameter', () => {
		const src =
			'Public Sub Mutate(ByRef value() As Long)\n' +
			'End Sub\n' +
			'Public Sub Caller()\n' +
			'    Dim x As Variant\n' +
			'    Mutate x\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), CODE, { span: 'x' });
	});

	it('does not flag an array passed to a Variant parameter (it boxes)', () => {
		const src =
			'Public Sub Mutate(ByRef value As Variant)\n' +
			'End Sub\n' +
			'Public Sub Caller()\n' +
			'    Dim arr(0 To 2) As Long\n' +
			'    Mutate arr\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});

	it('does not flag a Type value passed to a matching Type parameter', () => {
		const src =
			'Private Type T\n' +
			'    a As Long\n' +
			'End Type\n' +
			'Public Sub Mutate(ByRef value As T)\n' +
			'End Sub\n' +
			'Public Sub Caller()\n' +
			'    Dim u As T\n' +
			'    Mutate u\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});

	it('does not flag an array passed to a matching array parameter', () => {
		const src =
			'Public Sub Mutate(ByRef value() As Long)\n' +
			'End Sub\n' +
			'Public Sub Caller()\n' +
			'    Dim arr() As Long\n' +
			'    Mutate arr\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});

	it('does not flag an argument to a ParamArray parameter', () => {
		const src =
			'Public Sub TakesArgs(ParamArray items() As Variant)\n' +
			'End Sub\n' +
			'Public Sub Caller()\n' +
			'    Dim arr(0 To 2) As Long\n' +
			'    TakesArgs arr\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});

	it('does not flag an indexed array element passed to a scalar parameter', () => {
		const src =
			'Public Sub Mutate(ByRef value As Long)\n' +
			'End Sub\n' +
			'Public Sub Caller()\n' +
			'    Dim arr(0 To 2) As Long\n' +
			'    Mutate arr(0)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});

	it('stays quiet for an unresolved argument name', () => {
		const src =
			'Public Sub Mutate(ByRef value As Long)\n' +
			'End Sub\n' +
			'Public Sub Caller()\n' +
			'    Mutate undeclaredThing\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});

	it('defers to byref-argument-type-mismatch on element mismatch (no double-report)', () => {
		const src =
			'Public Sub Mutate(ByRef value As Long)\n' +
			'End Sub\n' +
			'Public Sub Caller()\n' +
			'    Dim arr(0 To 2) As Integer\n' +
			'    Mutate arr\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src);
		const shape = byCode(diagnostics, CODE).length;
		const byRef = byCode(diagnostics, 'byref-argument-type-mismatch').length;
		// Exactly one rule owns the slot; the two never both fire on `arr`.
		expect(shape + byRef).toBe(1);
	});

	it('ignores calls in inactive conditional-compilation branches', () => {
		const src =
			'#Const Enabled = False\n' +
			'Public Sub Mutate(ByRef value As Long)\n' +
			'End Sub\n' +
			'Public Sub Caller()\n' +
			'    Dim arr(0 To 2) As Long\n' +
			'#If Enabled Then\n' +
			'    Mutate arr\n' +
			'#End If\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});
});
