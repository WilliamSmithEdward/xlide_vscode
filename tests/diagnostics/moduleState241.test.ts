// Diagnostics tests: module-level and class-level state the runtime rules
// read (issue #241). A module variable nothing writes keeps its initial
// value everywhere. Each raising case was measured in Excel 16.0 (build
// 20326, 2026-09-30) and raises every time it runs; each quiet neighbour
// runs clean there.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { writtenNamesIn } from '../../src/analyzer/diagnostics/moduleState';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';
import { analyzeProjectModule } from './helpers';

function module(decls: string, main: string, rest = ''): string {
	return `Option Explicit\n${decls}\n\nFunction Main() As Variant\n    ${main}\nEnd Function\n${rest}`;
}

describe('writtenNamesIn', () => {
	it('takes targets, statement heads that write, and whole names passed to calls', () => {
		const names = writtenNamesIn([
			'Sub T()',
			'    a = 1',
			'    Set b = Nothing',
			'    ReDim c(3)',
			'    Fill d',
			'    Call Fill(e)',
			'    x = Take(f) + CLng(g)',
			'    If n = 1 Then h = 2',
			'    Debug.Print i',
			'    For j = 1 To 2: Next',
			'End Sub',
		].join('\n'));
		for (const name of ['a', 'b', 'c', 'd', 'e', 'f', 'h', 'j', 'x']) {
			expect(names.has(name), name).toBe(true);
		}
		// Read only: a library function's argument, a condition, a print.
		for (const name of ['g', 'n', 'i']) {
			expect(names.has(name), name).toBe(false);
		}
		// The module's own Trim is no library function, and a parameter is no write.
		const shadowed = writtenNamesIn('Sub T()\n    y = Trim(k)\nEnd Sub\nFunction Trim(x As String) As String\nEnd Function\nSub U(p)\nEnd Sub\n');
		expect(shadowed.has('k')).toBe(true);
		expect(shadowed.has('p')).toBe(false);
	});
});

describe('a module variable nothing writes', () => {
	it.each([
		['object-variable-not-set', 'Private m As Collection', 'Main = m.Count', '', 'm'],
		['object-variable-not-set', 'Private m As Collection', 'Use\n    Main = 1', '\nPrivate Sub Use()\n    Debug.Print m.Count\nEnd Sub\n', 'm'],
		['array-subscript-out-of-bounds', 'Private a() As Long', 'Main = a(0)', '', 'a'],
		['array-subscript-out-of-bounds', 'Private a() As Long', 'Main = UBound(a)', '', 'a'],
		['array-subscript-out-of-bounds', 'Private a(3) As Long', 'Main = a(5)', '', '5'],
		['division-by-zero', 'Private z As Long', 'Main = 10 / z', '', 'z'],
		['runtime-conversion-value', 'Private s As String', 'Main = CLng(s)', '', 's'],
	])('reports %s for %s', (code, decls, main, rest, span) => {
		const src = module(decls, main, rest);
		expectDiagnostic(src, byCode(analyzeModule(src), code), code, { span });
	});

	it('reports it though another procedure has a parameter of its name', () => {
		const src = module('Private m As Collection', 'Main = m.Count', '\nPrivate Sub Init(m)\n    Debug.Print m.Count\nEnd Sub\n');
		const hits = byCode(analyzeModule(src), 'object-variable-not-set');
		expect(hits).toHaveLength(1);
		expect(hits[0].message).toContain('never set anywhere in this module');
	});

	it('reads a class field the same way', () => {
		const cls = 'Option Explicit\nPrivate m As Collection\nPrivate a() As Long\n\nPublic Function Size() As Long\n    Size = m.Count + a(0)\nEnd Function\n';
		const diags = analyzeModule(cls, { moduleName: 'Class1', moduleKind: 'class' });
		expect(byCode(diags, 'object-variable-not-set')).toHaveLength(1);
		expect(byCode(diags, 'array-subscript-out-of-bounds')).toHaveLength(1);
	});

	it.each([
		['set in a Sub', 'Private m As Collection', 'Init\n    Main = m.Count', '\nPrivate Sub Init()\n    Set m = New Collection\nEnd Sub\n'],
		['As New', 'Private m As New Collection', 'Main = m.Count', ''],
		['ReDim\'d in a Sub', 'Private a() As Long', 'Init\n    Main = a(0)', '\nPrivate Sub Init()\n    ReDim a(3)\nEnd Sub\n'],
		['in range', 'Private a(3) As Long', 'Main = a(3)', ''],
		['assigned in a Sub', 'Private z As Long', 'Setup\n    Main = 10 / z', '\nPrivate Sub Setup()\n    z = 2\nEnd Sub\n'],
		['passed ByRef', 'Private s As String', 'Fill s\n    Main = CLng(s)', '\nPrivate Sub Fill(ByRef x As String)\n    x = "12"\nEnd Sub\n'],
		['hidden by a local', 'Private m As Collection', 'Dim m As New Collection\n    Main = m.Count', ''],
		['set in Main', 'Private m As Collection', 'Set m = New Collection\n    Main = m.Count', ''],
		['only compared with Nothing', 'Private m As Collection', 'If m Is Nothing Then Main = 1', ''],
		['a local dynamic array hides it', 'Private a(3) As Long', 'Dim a() As Long\n    ReDim a(9)\n    Main = a(5)', ''],
	])('stays quiet when %s', (_label, decls, main, rest) => {
		const diags = analyzeModule(module(decls, main, rest));
		for (const code of ['object-variable-not-set', 'array-subscript-out-of-bounds', 'division-by-zero', 'runtime-conversion-value']) {
			expect(byCode(diags, code), code).toHaveLength(0);
		}
	});

	it('takes the field Class_Initialize sets as set', () => {
		const cls = 'Option Explicit\nPrivate m As Collection\n\nPrivate Sub Class_Initialize()\n    Set m = New Collection\nEnd Sub\n\nPublic Function Size() As Long\n    Size = m.Count\nEnd Function\n';
		expect(byCode(analyzeModule(cls, { moduleName: 'Class1', moduleKind: 'class' }), 'object-variable-not-set')).toHaveLength(0);
	});
});

describe('a Public variable', () => {
	const main = module('Public g As Collection', 'Main = g.Count');

	it('is reported when no module of the project writes it', () => {
		const diags = analyzeProjectModule(main, [{ moduleName: 'Module1', source: main }, { moduleName: 'Module2', source: 'Option Explicit\n' }], 'Module1');
		expectDiagnostic(main, byCode(diags, 'object-variable-not-set'), 'object-variable-not-set', { span: 'g', message: 'anywhere in the project' });
	});

	it('stays quiet when another module sets it, or without the project', () => {
		const setter = 'Option Explicit\n\nPublic Sub Init()\n    Set g = New Collection\nEnd Sub\n';
		const diags = analyzeProjectModule(main, [{ moduleName: 'Module1', source: main }, { moduleName: 'Module2', source: setter }], 'Module1');
		expect(byCode(diags, 'object-variable-not-set')).toHaveLength(0);
		expect(byCode(analyzeModule(main), 'object-variable-not-set')).toHaveLength(0);
	});
});
