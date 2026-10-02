// Diagnostics tests: an object of the wrong class reaching a typed variable
// through an Object variable, an argument or a Collection item (issue #246).
// Each raising case was measured in Excel 16.0 (build 20326, 2026-09-30);
// each quiet neighbour runs clean there.

import { describe, it, expect } from 'vitest';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';
import { analyzeProjectModule } from './helpers';

const CLASSES = [
	{ moduleName: 'IArea1', type: 'class' as const, source: 'Option Explicit\nPublic Function Area() As Double\nEnd Function\n' },
	{ moduleName: 'Round1', type: 'class' as const, source: 'Option Explicit\nImplements IArea1\nPrivate Function IArea1_Area() As Double\n    IArea1_Area = 3\nEnd Function\nPublic Function Radius() As Double\n    Radius = 1\nEnd Function\n' },
	{ moduleName: 'Flat1', type: 'class' as const, source: 'Option Explicit\nPublic Function Area() As Double\n    Area = 2\nEnd Function\n' },
];
const HELPERS = 'Private Sub Fill(o As Object)\nEnd Sub\n\nPrivate Function TakeRound1(ByVal c As Round1) As Long\nEnd Function\n\nPrivate Function TakeRound1Ref(c As Round1) As Long\nEnd Function\n\nPrivate Sub UseRound1(c As Round1)\nEnd Sub\n';

function analyze(...lines: string[]) {
	const src = `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n\n${HELPERS}`;
	return { src, diags: analyzeProjectModule(src, [{ moduleName: 'Module1', source: src }, ...CLASSES], 'Module1') };
}

describe('through an Object variable', () => {
	it('reports a Set from an Object that holds another class', () => {
		const { src, diags } = analyze('Dim o As Object, c As Round1', 'Set o = New Flat1', 'Set c = o');
		expectDiagnostic(src, byCode(diags, 'assignment-object-type-mismatch'), 'assignment-object-type-mismatch', { message: ["'o', which holds a Flat1 here", "'13'"] });
	});

	it('follows the class from one Object to another', () => {
		const { src, diags } = analyze('Dim a As Object, b As Object, c As Round1', 'Set a = New Flat1', 'Set b = a', 'Set c = b');
		expectDiagnostic(src, byCode(diags, 'assignment-object-type-mismatch'), 'assignment-object-type-mismatch', { message: "'b', which holds a Flat1 here" });
	});

	it('forgets the class after Nothing, a ByRef pass or a label a GoTo names', () => {
		for (const lines of [
			['Dim o As Object, c As Round1', 'Set o = New Flat1', 'Set o = Nothing', 'Set c = o'],
			['Dim o As Object, c As Round1', 'Set o = New Flat1', 'Fill o', 'Set c = o'],
			['Dim o As Object, c As Round1', 'Set o = New Flat1', 'Again:', 'Set c = o', 'If c Is Nothing Then GoTo Again'],
		]) {
			expect(byCode(analyze(...lines).diags, 'assignment-object-type-mismatch'), lines.join('; ')).toHaveLength(0);
		}
	});

	it('stays quiet when it holds the class, or an interface it implements', () => {
		expect(byCode(analyze('Dim o As Object, c As Round1', 'Set o = New Round1', 'Set c = o').diags, 'assignment-object-type-mismatch')).toHaveLength(0);
		expect(byCode(analyze('Dim s As IArea1, c As Round1', 'Set s = New Round1', 'Set c = s').diags, 'assignment-object-type-mismatch')).toHaveLength(0);
		expect(byCode(analyze('Dim o As Object, c As Round1', 'Set o = New Flat1', 'Set o = New Round1', 'Set c = o').diags, 'assignment-object-type-mismatch')).toHaveLength(0);
	});
});

describe('through an argument', () => {
	it.each([
		['Main = TakeRound1Ref(p)'],
		['Main = TakeRound1(p)'],
		['UseRound1 p'],
	])('reports a variable that holds a Flat1 in %s', (call) => {
		const { src, diags } = analyze('Dim p As Flat1', 'Set p = New Flat1', call);
		expectDiagnostic(src, byCode(diags, 'argument-type-mismatch'), 'argument-type-mismatch', { span: 'p', message: ["'p', which holds a Flat1 here", "'13'"] });
	});

	it('stays quiet on a variable still Nothing, which passes', () => {
		expect(byCode(analyze('Dim p As Flat1', 'Main = TakeRound1Ref(p)').diags, 'argument-type-mismatch')).toHaveLength(0);
	});
});

describe('through a Collection item', () => {
	it('reports a member the item lacks, and a For Each into a class it is not', () => {
		const member = analyze('Dim c As New Collection', 'c.Add New Flat1', 'Main = c(1).Radius()');
		expectDiagnostic(member.src, byCode(member.diags, 'runtime-member-not-found'), 'runtime-member-not-found', { span: 'c(1)', message: ['holds a Flat1', "'438'"] });
		const loop = analyze('Dim c As New Collection, x As Round1', 'c.Add New Flat1', 'For Each x In c', 'Next');
		expectDiagnostic(loop.src, byCode(loop.diags, 'assignment-object-type-mismatch'), 'assignment-object-type-mismatch', { span: 'c', message: ['item 1 is a Flat1', "'13'"] });
	});

	it('picks the item by a literal index when the classes differ', () => {
		const second = analyze('Dim c As New Collection', 'c.Add New Round1', 'c.Add New Flat1', 'Main = c(2).Radius()');
		expectDiagnostic(second.src, byCode(second.diags, 'runtime-member-not-found'), 'runtime-member-not-found', { span: 'c(2)' });
		for (const lines of [
			['Dim c As New Collection', 'c.Add New Round1', 'c.Add New Flat1', 'Main = c(1).Radius()'],
			['Dim c As New Collection, k As Long', 'c.Add New Flat1', 'c.Add New Round1', 'k = Len("ab")', 'Main = c(k).Radius()'],
		]) {
			expect(byCode(analyze(...lines).diags, 'runtime-member-not-found'), lines.join('; ')).toHaveLength(0);
		}
	});

	it('stays quiet on items of the right class, and once the Collection changes unseen', () => {
		for (const lines of [
			['Dim c As New Collection', 'c.Add New Round1', 'Main = c(1).Radius()'],
			['Dim c As New Collection, x As Round1', 'c.Add New Round1', 'For Each x In c', 'Next'],
			['Dim c As New Collection', 'c.Add New Flat1', 'c.Remove 1', 'c.Add New Round1', 'Main = c(1).Radius()'],
			['Dim c As New Collection, o As Object', 'Set o = New Round1', 'c.Add o', 'Main = c(1).Radius()'],
		]) {
			const { diags } = analyze(...lines);
			expect(byCode(diags, 'runtime-member-not-found').length + byCode(diags, 'assignment-object-type-mismatch').length, lines.join('; ')).toBe(0);
		}
	});

	// Issue #356, measured in Excel 16.0 (build 20326, 2026-10-01).
	it('follows Before and After, and forgets the order where they are not a whole number', () => {
		for (const lines of [
			['Dim c As New Collection', 'c.Add New Flat1', 'c.Add New Round1, Before:=1', 'Main = c(1).Radius()'],
			['Dim c As New Collection', 'c.Add New Flat1', 'c.Add New Round1, , 1', 'Main = c(1).Radius()'],
			['Dim c As New Collection', 'c.Add New Round1', 'c.Add New Flat1', 'c.Add New Round1, After:=1', 'Main = c(2).Radius()'],
			['Dim c As New Collection', 'c.Add New Flat1, "f"', 'c.Add New Round1, "r", "f"', 'Main = c(1).Radius()'],
			['Dim c As New Collection', 'c.Add New Round1', 'c.Add New Flat1', 'c.Add New Round1, After:=0 + 1', 'Main = c(2).Radius()'],
			['Dim c As New Collection, k As Long', 'c.Add New Round1', 'c.Add New Flat1', 'k = 2', 'c.Add New Round1, After:=k', 'Main = c(3).Radius()'],
		]) {
			expect(byCode(analyze(...lines).diags, 'runtime-member-not-found'), lines.join('; ')).toHaveLength(0);
		}
		for (const [lines, span] of [
			[['Dim c As New Collection', 'c.Add New Round1', 'c.Add New Flat1, Before:=1', 'Main = c(1).Radius()'], 'c(1)'],
			[['Dim c As New Collection', 'c.Add New Round1', 'c.Add New Round1', 'c.Add New Flat1, After:=1', 'Main = c(2).Radius()'], 'c(2)'],
		] as const) {
			const found = analyze(...lines);
			expectDiagnostic(found.src, byCode(found.diags, 'runtime-member-not-found'), 'runtime-member-not-found', { span });
		}
	});

	it('judges only the first item of a For Each that may leave', () => {
		for (const exit of ['Exit For', 'If x.Radius() >= 0 Then Exit For', 'Exit Function', 'GoTo Done']) {
			const lines = ['Dim c As New Collection, x As Round1', 'c.Add New Round1', 'c.Add New Flat1', 'For Each x In c', `    ${exit}`, 'Next', 'Done:'];
			expect(byCode(analyze(...lines).diags, 'assignment-object-type-mismatch'), exit).toHaveLength(0);
		}
		const first = analyze('Dim c As New Collection, x As Round1', 'c.Add New Flat1', 'c.Add New Round1', 'For Each x In c', '    Exit For', 'Next');
		expectDiagnostic(first.src, byCode(first.diags, 'assignment-object-type-mismatch'), 'assignment-object-type-mismatch', { span: 'c', message: 'item 1 is a Flat1' });
	});
});
