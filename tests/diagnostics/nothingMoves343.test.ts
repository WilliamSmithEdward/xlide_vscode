// Diagnostics tests: Nothing once it moves to another variable (issue
// #343). Each raising sample was measured through pyVBAharness on
// 2026-10-02 in Excel 16.0 (build 20326); each quiet one runs there.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode } from '../helpers/diagnostics';

const NOT_SET = 'object-variable-not-set';

const HELPERS = [
	'Function GetNothing() As Collection', '    Set GetNothing = Nothing', 'End Function',
	'Function TakeC(ByVal c As Collection) As Variant', '    TakeC = c.Count', 'End Function',
	'Function Guarded(ByVal c As Collection) As Variant', '    If c Is Nothing Then Exit Function', '    Guarded = c.Count', 'End Function',
	'Sub Fill(c As Collection)', '    Set c = New Collection', 'End Sub',
].join('\n');

function wrap(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n${HELPERS}\n`;
}

describe('Nothing once it moves (issue #343)', () => {
	it('follows a copy, a Variant set to Nothing and a Function that returns Nothing', () => {
		const states = [
			['Dim o As Collection, p As Collection', 'Set o = p'],
			['Dim p As Collection, o As Collection', 'Set o = p'],
			['Dim o As Variant', 'Set o = Nothing'],
			['Dim o', 'Set o = Nothing'],
			['Dim o As Collection', 'Set o = GetNothing()'],
		];
		for (const state of states) {
			for (const use of ['Main = o.Count', 'o.Add 1', 'Main = IIf(o Is Nothing, 0, o.Count)']) {
				const hits = byCode(analyzeModule(wrap(...state, use)), NOT_SET);
				expect(hits, `${state.join(' / ')} / ${use}`).toHaveLength(1);
			}
		}
	});

	it('reports Nothing passed to a procedure that reads a member of it first', () => {
		for (const use of ['Main = TakeC(o)', 'Call TakeC(o)', 'TakeC o']) {
			const hits = byCode(analyzeModule(wrap('Dim o As Collection', use)), NOT_SET);
			expect(hits, use).toHaveLength(1);
			expect(hits[0].message, use).toContain("'c.Count'");
		}
	});

	it('passes an Object ByRef to a Collection parameter, and back', () => {
		const src = 'Option Explicit\nFunction Main() As Variant\n    Dim o As Object, c As Collection\n    Set o = New Collection\n    Set c = o\n'
			+ '    Main = TakeR(o) + Len(TakeO(c))\nEnd Function\n'
			+ 'Function TakeR(c As Collection) As Variant\n    TakeR = c.Count\nEnd Function\n'
			+ 'Function TakeO(o As Object) As Variant\n    TakeO = TypeName(o)\nEnd Function\n';
		expect(byCode(analyzeModule(src), 'byref-argument-type-mismatch')).toHaveLength(0);
	});

	it('stays quiet where the object is there, or the callee checks or sets it', () => {
		const bodies = [
			['Dim o As Collection, p As Collection', 'Set p = New Collection', 'Set o = p', 'Set p = Nothing', 'Main = o.Count'],
			['Dim o As Collection', 'Main = Guarded(o)'],
			['Dim o As Collection', 'Fill o', 'Main = TakeC(o)'],
			['Dim o As Variant', 'Set o = Nothing', 'o = 5', 'Main = o'],
			['Dim o As Variant', 'Set o = Nothing', 'Set o = New Collection', 'Main = o.Count'],
			['Dim o As Variant', 'Main = TypeName(o)'],
		];
		for (const body of bodies) {
			expect(byCode(analyzeModule(wrap(...body)), NOT_SET), body.join(' / ')).toHaveLength(0);
		}
	});
});
