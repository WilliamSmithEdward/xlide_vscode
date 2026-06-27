// Diagnostics tests: typeof-is-always-false rule (expression-binder Slice 4).
//
// `TypeOf x Is Y` is flagged only when x's declared object type can never be a Y.
// Verified against MS-VBAL 5.6 (TypeOf...Is) plus the shared object tables.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, spanText } from '../helpers/diagnostics';
import { projectClassMembers } from './helpers';

const CODE = 'typeof-is-always-false';

function hits(src: string) {
	return byCode(analyzeModule(src), CODE);
}

function wrap(body: string): string {
	return `Sub T()\n${body}\nEnd Sub\n`;
}

describe('analyzeModule - typeof-is-always-false', () => {
	it('flags an If condition where the operand can never be the target type', () => {
		const src = wrap('    Dim wb As Workbook\n    If TypeOf wb Is Worksheet Then\n        Debug.Print 1\n    End If');
		const found = hits(src);
		expect(found).toHaveLength(1);
		expect(found[0].severity).toBe('warning');
		expect(spanText(src, found[0])).toBe('TypeOf wb Is Worksheet');
	});

	it('flags a TypeOf used outside an If condition (assignment RHS)', () => {
		const src = wrap('    Dim x As Range\n    Dim b As Boolean\n    b = (TypeOf x Is Worksheet)');
		expect(hits(src)).toHaveLength(1);
	});

	it('stays quiet when operand and target are the same type', () => {
		expect(hits(wrap('    Dim ws As Worksheet\n    If TypeOf ws Is Worksheet Then\n    End If'))).toHaveLength(0);
	});

	it('stays quiet for an Object or Variant operand (could hold anything)', () => {
		expect(hits(wrap('    Dim o As Object\n    If TypeOf o Is Worksheet Then\n    End If'))).toHaveLength(0);
		expect(hits(wrap('    Dim v As Variant\n    If TypeOf v Is Worksheet Then\n    End If'))).toHaveLength(0);
	});

	it('stays quiet for `Is Object` (always True, not False)', () => {
		expect(hits(wrap('    Dim wb As Workbook\n    If TypeOf wb Is Object Then\n    End If'))).toHaveLength(0);
	});

	it('stays quiet when the operand type is unknown', () => {
		expect(hits(wrap('    Dim x As FooUnknownType\n    If TypeOf x Is Worksheet Then\n    End If'))).toHaveLength(0);
	});

	it('stays quiet for a scalar operand', () => {
		expect(hits(wrap('    Dim n As Long\n    If TypeOf n Is Worksheet Then\n    End If'))).toHaveLength(0);
	});

	it('stays quiet for a non-identifier operand (not yet modeled)', () => {
		// Member-access operand: the operand type is not resolved in this slice.
		expect(hits(wrap('    Dim wb As Workbook\n    If TypeOf wb.ActiveSheet Is Workbook Then\n    End If'))).toHaveLength(0);
	});

	it('stays quiet for a project interface operand or a class assignable to it', () => {
		// Concrete-operand gate: Person is an interface Class1 Implements, so an
		// interface-typed operand could hold a Class1 (a), and Class1 is-a Person
		// (b) - both mutually assignable, so neither TypeOf test is always-False.
		const project = projectClassMembers([
			{ moduleName: 'Person', moduleKind: 'class', source: 'Public Sub Save()\nEnd Sub\n' },
			{ moduleName: 'Class1', moduleKind: 'class', source: 'Implements Person\n' },
		]);
		const interfaceOperand =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    If TypeOf p Is Class1 Then\n' +
			'    End If\n' +
			'End Sub\n';
		const implementorOperand =
			'Public Sub T()\n' +
			'    Dim c As Class1\n' +
			'    If TypeOf c Is Person Then\n' +
			'    End If\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(interfaceOperand, { projectClassMembers: project }), CODE)).toHaveLength(0);
		expect(byCode(analyzeModule(implementorOperand, { projectClassMembers: project }), CODE)).toHaveLength(0);
	});
});

describe('analyzeModule - is-operator-non-object', () => {
	const IS = 'is-operator-non-object';

	it('flags Is with a Long-declared (scalar) operand', () => {
		const src = 'Sub T()\n\tDim n As Long\n\tIf n Is Nothing Then\n\tEnd If\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), IS);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('n');
		expect(hits[0].severity).toBe('error');
	});

	it('flags Is with a String-declared operand', () => {
		const src = 'Sub T()\n\tDim s As String\n\tIf s Is Nothing Then\n\tEnd If\nEnd Sub\n';
		expect(byCode(analyzeModule(src), IS)).toHaveLength(1);
	});

	it('flags Is with two scalar operands (fires on the first)', () => {
		const src = 'Sub T()\n\tDim a As Long, b As Long\n\tIf a Is b Then\n\tEnd If\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), IS);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('a');
	});

	it('flags Is with scalar value literals', () => {
		const src = 'Sub T()\n\tIf 1 Is 1 Then\n\tEnd If\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), IS);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('1');
	});

	it('does not flag Is with an Object-declared operand', () => {
		const src = 'Sub T()\n\tDim o As Object\n\tIf o Is Nothing Then\n\tEnd If\nEnd Sub\n';
		expect(byCode(analyzeModule(src), IS)).toHaveLength(0);
	});

	it('does not flag Is with a Variant operand (may hold an object)', () => {
		const src = 'Sub T()\n\tDim v As Variant\n\tIf v Is Nothing Then\n\tEnd If\nEnd Sub\n';
		expect(byCode(analyzeModule(src), IS)).toHaveLength(0);
	});

	it('does not flag Is with a class-typed operand', () => {
		const src = 'Sub T()\n\tDim c As Collection\n\tIf c Is Nothing Then\n\tEnd If\nEnd Sub\n';
		expect(byCode(analyzeModule(src), IS)).toHaveLength(0);
	});

	it('does not flag Is with an undeclared operand', () => {
		const src = 'Sub T()\n\tIf x Is Nothing Then\n\tEnd If\nEnd Sub\n';
		expect(byCode(analyzeModule(src), IS)).toHaveLength(0);
	});

	it('does not flag the valid Nothing Is Nothing idiom', () => {
		const src = 'Sub T()\n\tIf Nothing Is Nothing Then\n\tEnd If\nEnd Sub\n';
		expect(byCode(analyzeModule(src), IS)).toHaveLength(0);
	});

	it('does not touch TypeOf ... Is (owned by typeof-is-always-false)', () => {
		const src = 'Sub T()\n\tDim o As Object\n\tIf TypeOf o Is Collection Then\n\tEnd If\nEnd Sub\n';
		expect(byCode(analyzeModule(src), IS)).toHaveLength(0);
	});

	it('does not flag a vb-prefixed class type (vbLong is an object, not a scalar)', () => {
		const src = 'Sub T()\n\tDim x As vbLong\n\tIf x Is Nothing Then\n\tEnd If\nEnd Sub\n';
		expect(byCode(analyzeModule(src), IS)).toHaveLength(0);
	});

	it('does not flag a vb-prefixed class type with a scalar-word suffix (vbString)', () => {
		const src = 'Sub T()\n\tDim x As vbString\n\tIf x Is Nothing Then\n\tEnd If\nEnd Sub\n';
		expect(byCode(analyzeModule(src), IS)).toHaveLength(0);
	});

	it('flags a bare array variable (arrays are not object references)', () => {
		const src = 'Sub T()\n\tDim arr() As Long\n\tIf arr Is Nothing Then\n\tEnd If\nEnd Sub\n';
		expect(byCode(analyzeModule(src), IS)).toHaveLength(1);
	});
});
