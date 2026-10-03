// Diagnostics tests: a member of a class instance the procedure makes, used
// where what it holds cannot serve (issue #414). Measured on 2026-10-02 in
// Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeProjectModule } from './helpers';

const FIELD_OBJ = 'Public M As Object\n';
const FIELD_COLL = 'Public M As Collection\n';
const FIELD_VAR = 'Public M As Variant\n';
const FUNC_NOTHING = 'Public Function M() As Object\n    Set M = Nothing\nEnd Function\n';
const GET_VAR1 = 'Public Property Get M() As Variant\n    M = 1\nEnd Property\n';
const SUB = 'Public Sub M()\nEnd Sub\n';
const GET_LONG = 'Public Property Get M() As Long\n    M = 1\nEnd Property\n';
const SET_ONLY = 'Public Property Set M(ByVal v As Object)\nEnd Property\n';
const FIELD_STR = 'Public M As String\n';
const LET_ONLY = 'Public Property Let M(ByVal v As Variant)\nEnd Property\n';
const LET_INDEXED = 'Public Property Let M(ByVal i As Long, ByVal v As Variant)\nEnd Property\n';
const INIT_OBJ = 'Public M As Object\nPrivate Sub Class_Initialize()\n    Set M = New Collection\nEnd Sub\n';

function raised(member: string, decl: string, body: string, extra = ''): string[] {
	const setup = decl.startsWith('New') ? `Dim c As ${decl}` : `Dim c As ${decl}\n    Set c = New Class1`;
	const src = `Option Explicit\n${extra}Function Main() As Variant\n    ${setup}\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeProjectModule(src, [{ moduleName: 'Module1', source: src }, { moduleName: 'Class1', source: `Option Explicit\n${member}`, moduleKind: 'class' }], 'Module1')
		.filter((d) => d.severity === 'error')
		.map((d) => /Run-time error '(-?\d+)'/.exec(d.message)?.[1] ?? d.code);
}

describe('a class instance\'s member, from what the class shows (issue #414)', () => {
	it('reports what it holds', () => {
		const cases: Array<[string, string, string, string]> = [
			[FIELD_COLL, 'New Class1', 'c.M.Add 1', '91'],
			[FIELD_OBJ, 'New Class1', 'Main = c.M & "x"', '91'],
			[FIELD_OBJ, 'New Class1', 'Main = c.M(1)', '91'],
			[FIELD_OBJ, 'New Class1', 'Main = c.M.Count', '91'],
			[FIELD_OBJ, 'Class1', 'Main = c.M.Count', '91'],
			[FIELD_COLL, 'Object', 'Main = c.M.Count', '91'],
			[FUNC_NOTHING, 'New Class1', 'Main = c.M', '91'],
			[GET_VAR1, 'New Class1', 'c.M(1) = 2', '13'],
			[GET_VAR1, 'New Class1', 'Main = c.M(1)', '13'],
			[GET_VAR1, 'New Class1', 'Dim o As Object\n    Set o = c.M', '424'],
			[FIELD_VAR, 'New Class1', 'Main = c.M.Count', '424'],
		];
		for (const [member, decl, body, error] of cases) {
			expect(raised(member, decl, body), `${member} ${body}`).toEqual([error]);
		}
	});

	it('reports late-bound uses the class refuses', () => {
		const cases: Array<[string, string, string, string]> = [
			[SUB, 'Object', 'c.M = 5', '450'],
			[SUB, 'Variant', 'c.M = 5', '450'],
			[SUB, 'Object', 'Main = c.M(1)', '451'],
			[GET_LONG, 'Object', 'c.M = 1', '451'],
			[SET_ONLY, 'Object', 'Main = c.M', '450'],
			[FIELD_STR, 'Object', 'c.M.Add 1', '424'],
		];
		for (const [member, decl, body, error] of cases) {
			expect(raised(member, decl, body), `${member} ${body}`).toEqual([error]);
		}
	});

	it('reports an element of a Let-only property assigned, a compile error', () => {
		expect(raised(LET_ONLY, 'New Class1', 'c.M(1) = 2')).toEqual(['invalid-property-use']);
		expect(raised(LET_INDEXED, 'New Class1', 'c.M(1) = 2')).toEqual([]);
	});

	it('leaves alone what runs', () => {
		const fill = 'Private Sub Fill(ByVal k As Object)\n    Set k.M = New Collection\nEnd Sub\n';
		const quiet: Array<[string, string, string, string?]> = [
			[INIT_OBJ, 'New Class1', 'Main = c.M.Count'],
			['Public Function M() As Object\n    Set M = New Collection\nEnd Function\n', 'New Class1', 'Main = c.M.Count'],
			[FIELD_OBJ, 'New Class1', 'Fill c\n    Main = c.M.Count', fill],
			[FIELD_OBJ, 'New Class1', 'Main = c.M Is Nothing'],
			[FIELD_OBJ, 'New Class1', 'If c.M Is Nothing Then Main = 2'],
			[FIELD_COLL, 'New Class1', 'Set c.M = New Collection\n    c.M.Add 1'],
			[FIELD_OBJ, 'New Class1', 'Dim o As Object\n    Set o = c.M\n    Main = o Is Nothing'],
			[GET_LONG, 'Object', 'Main = c.M'],
			[SUB, 'Object', 'c.M'],
			[SET_ONLY, 'Object', 'Set c.M = New Collection'],
			[LET_ONLY, 'Object', 'c.M = 5'],
			[GET_VAR1, 'Object', 'Main = c.M'],
			[GET_VAR1, 'New Class1', 'Main = c.M'],
		];
		for (const [member, decl, body, extra] of quiet) {
			expect(raised(member, decl, body, extra), `${member} ${body}`).toEqual([]);
		}
	});
});
