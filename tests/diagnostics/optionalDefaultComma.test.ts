// Diagnostics tests: a parameter whose Optional default is a string holding a
// comma, a quote or a bracket (`Optional ByVal d As String = ","`) must not
// be counted as more than one parameter when a member signature is read back
// for its argument counts, early bound, by name or late bound.

import { describe, it, expect } from 'vitest';
import { analyzeProjectModule } from './helpers';

const CLASS = 'Option Explicit\n'
	+ 'Public Function Comma(Optional ByVal d As String = ",") As String\nEnd Function\n'
	+ 'Public Function List(Optional ByVal d As String = "a,b,c") As String\nEnd Function\n'
	+ 'Public Function Quoted(Optional ByVal d As String = """,""") As String\nEnd Function\n'
	+ 'Public Function Paren(Optional ByVal n As Long = (1 + 2), Optional ByVal m As Long = 3) As Long\nEnd Function\n'
	+ 'Public Function Closer(Optional ByVal d As String = ")", Optional ByVal e As Long = 1) As Long\nEnd Function\n'
	+ 'Public Function Bracket(Optional ByVal d As String = "]", Optional ByVal e As Long = 1) As Long\nEnd Function\n'
	+ 'Public Function Needs(ByVal a As Long, Optional ByVal d As String = ",") As String\nEnd Function\n'
	+ 'Public Function Many(ByVal a As Long, ParamArray v() As Variant) As Long\nEnd Function\n'
	+ 'Public Function Mixed(Optional ByVal d As String = ",", ByVal a As Long) As Long\nEnd Function\n'
	+ 'Public Function Arr(a() As Long, ByVal b As Long) As Long\nEnd Function\n';
const MODULE_TAIL = 'Public Function Join2(Optional ByVal d As String = ",") As String\nEnd Function\n';

function errors(line: string, declare = 'Dim c As New Class1'): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    ${declare}\n    ${line}\nEnd Function\n${MODULE_TAIL}`;
	return analyzeProjectModule(src, [
		{ moduleName: 'Module1', source: src },
		{ moduleName: 'Class1', source: CLASS, type: 'class' },
	], 'Module1', { host: 'excel' }).filter((diag) => diag.severity === 'error').map((diag) => `${diag.code}: ${diag.message}`);
}

describe('an Optional default holding a comma, quote or bracket', () => {
	it('is one parameter when a class member is read with no argument', () => {
		for (const line of ['Main = c.Comma', 'Main = c.List', 'Main = c.Quoted', 'Main = c.Paren', 'Main = c.Closer', 'Main = c.Bracket']) {
			expect(errors(line), line).toEqual([]);
		}
	});

	it('is one parameter when the member is called by name', () => {
		for (const line of [
			'Main = CallByName(c, "Comma", VbMethod)', 'Main = CallByName(c, "List", VbMethod)', 'Main = CallByName(c, "Quoted", VbMethod)',
			'Main = CallByName(c, "Paren", VbMethod, 1, 2)', 'Main = CallByName(c, "Closer", VbMethod, "x", 2)', 'Main = CallByName(c, "Bracket", VbMethod, "x", 2)',
			'Main = CallByName(c, "Many", VbMethod, 1)', 'Main = CallByName(c, "Many", VbMethod, 1, 2, 3)',
			'Main = Application.Run("Join2")', 'Main = Application.Run("Join2", ";")',
		]) {
			expect(errors(line), line).toEqual([]);
		}
	});

	it('still reports a required parameter that is missing or an argument too many', () => {
		for (const line of ['Main = c.Needs', 'Main = c.Many', 'Main = c.Mixed']) {
			expect(errors(line), line).toEqual([expect.stringMatching(/^argument-count: .*needs an argument, and is read here with none/)]);
		}
		expect(errors('Main = c.Arr')).toEqual([expect.stringMatching(/^argument-count: .*needs 2 arguments, and is read here with none/)]);
		for (const [line, error] of [
			['Main = CallByName(c, "Needs", VbMethod)', '449'], ['Main = CallByName(c, "Many", VbMethod)', '449'], ['Main = CallByName(c, "Mixed", VbMethod)', '449'],
			['Main = CallByName(c, "Comma", VbMethod, ",", ";")', '450'], ['Main = CallByName(c, "Bracket", VbMethod, "x", 2, 3)', '450'],
			['Main = Application.Run("Join2", ",", ";")', '450'],
		]) {
			expect(errors(line), line).toEqual([expect.stringMatching(new RegExp(`^runtime-member-not-found: .*'${error}'`))]);
		}
	});

	it('is one parameter when a late-bound call is checked against it', () => {
		const late = (line: string) => errors(line, 'Dim o As Object\n    Set o = New Class1');
		for (const line of ['Main = o.Comma', 'Main = o.Comma(",")', 'Main = o.Needs(1, ",")', 'Main = o.Needs(d:=",", a:=1)', 'Main = o.Closer("x", 2)', 'Main = o.Bracket("x", 2)']) {
			expect(late(line), line).toEqual([]);
		}
		for (const [line, error] of [
			['Main = o.Needs()', '449'], ['Main = o.Needs(1, ",", 3)', '450'], ['Main = o.Comma(e:=1)', '448'], ['Main = o.Closer("x", 2, 3)', '450'],
		]) {
			expect(late(line), line).toEqual([expect.stringMatching(new RegExp(`^runtime-member-not-found: .*'${error}'`))]);
		}
	});
});
