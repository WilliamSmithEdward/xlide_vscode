// Diagnostics tests: member access on a project class (issue #224). Each case
// was measured in 64-bit Excel 16.0 (build 20326, 2026-09-30).

import { describe, it, expect } from 'vitest';
import { analyzeProjectModule } from './helpers';

const CLASS1 = [
	'Public Field As Long',
	'Public Name1 As String',
	'Public Coll As Collection',
	'Public V As Variant',
	'Public Property Get RO() As Long',
	'End Property',
	'Public Property Get Idx(ByVal i As Long) As Long',
	'End Property',
	'Public Property Let Idx(ByVal i As Long, ByVal v As Long)',
	'End Property',
	'Public Property Get OptIdx(Optional ByVal i As Long = 1) As Long',
	'End Property',
	'Public Property Let WO(ByVal v As Long)',
	'End Property',
	'Public Function Calc(ByVal a As Long, Optional ByVal b As Long = 1) As Long',
	'End Function',
	'Public Sub Go(ByVal a As Long)',
	'End Sub',
	'Private Sub Hidden()',
	'End Sub',
].join('\n');

function errors(...lines: string[]): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    Dim c As New Class1\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
	return analyzeProjectModule(src, [
		{ moduleName: 'Module1', source: src },
		{ moduleName: 'Class1', source: CLASS1, type: 'class' },
	], 'Module1').filter((d) => d.severity === 'error').map((d) => `${d.code}: ${d.message}`);
}

describe('class members used in a form the VBE refuses (issue #224)', () => {
	it.each([
		['Main = c.Idx', 'Argument not optional'],
		['c.Idx = 5', 'Argument not optional'],
		['Main = c.Idx + 1', 'Argument not optional'],
		['Debug.Print c.Idx', 'Argument not optional'],
		['Main = c.Field(1)', 'Wrong number of arguments'],
		['Main = c.Name1(1)', 'Wrong number of arguments'],
		['c.Field(1) = 5', "Can't assign to read-only property"],
	])('flags %s', (line, text) => {
		const hits = errors(line);
		expect(hits).toHaveLength(1);
		expect(hits[0]).toMatch(/^argument-count: /);
		expect(hits[0]).toContain(text);
	});

	it.each([
		'Main = c.Idx(2)',
		'c.Idx(2) = 5',
		'Main = c.OptIdx',
		'Main = c.Field',
		'Main = c.Field()',
		'c.V = Array(1, 2): Main = c.V(1)',
		'Main = 1: If False Then Main = c.Coll(1)',
		'Main = c.RO',
	])('stays quiet on %s', (line) => {
		expect(errors(line)).toHaveLength(0);
	});
});

describe('parentheses around a member call statement (issue #224)', () => {
	it.each([
		'c.Calc (1, 2)',
		'c.Calc(1, 2)',
		'c.Go (1, 2)',
		'With c: .Calc (1, 2): End With',
	])('flags %s as a Syntax error', (line) => {
		const hits = errors(line, 'Main = 1');
		expect(hits.some((hit) => hit.startsWith('call-statement-multi-arg-parens: '))).toBe(true);
	});

	it.each(['c.Calc (1)', 'c.Go (1)', 'Call c.Calc(1, 2)', 'c.Calc 1, 2'])('stays quiet on %s', (line) => {
		expect(errors(line, 'Main = 1')).toHaveLength(0);
	});
});

describe('a class held in a late-bound variable (issue #224)', () => {
	it.each([
		[['Dim o As Object', 'Set o = c', 'o.Hidden'], "'438'"],
		[['Dim o As Object', 'Set o = c', 'o.Nope'], "'438'"],
		[['Dim o As Variant', 'Set o = c', 'o.Hidden'], "'438'"],
		[['Dim o As Variant', 'Set o = c', 'o.RO = 5'], "'451'"],
		[['Dim o As Object', 'Set o = c', 'o.RO = 5'], "'451'"],
		[['Dim o As Object', 'Set o = c', 'Main = o.WO'], "'450'"],
	])('flags %j', (lines, error) => {
		const hits = errors(...lines, 'Main = 1');
		expect(hits).toHaveLength(1);
		expect(hits[0]).toMatch(/^runtime-member-not-found: /);
		expect(hits[0]).toContain(error);
		expect(hits[0]).not.toContain("'91'");
	});

	it('names 91 as well where the variable it came from may be Nothing', () => {
		const hits = errors('Dim c2 As Class1', 'Dim o As Object', 'Set o = c2', 'o.Nope', 'Main = 1');
		expect(hits).toHaveLength(1);
		expect(hits[0]).toContain("'438'");
		expect(hits[0]).toContain("'91'");
	});

	it.each([
		['Dim o As Object', 'Set o = c', 'Main = o.RO'],
		['Dim o As Object', 'Set o = c', 'o.Go 1'],
		['Dim o As Object', 'Set o = c', 'o.WO = 5'],
		['Dim o As Object', 'Set o = c', 'o.Field = 3', 'Main = o.Field'],
	])('stays quiet on %j', (...lines) => {
		expect(errors(...lines, 'Main = 1')).toHaveLength(0);
	});
});
