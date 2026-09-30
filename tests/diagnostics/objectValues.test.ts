// Diagnostics tests: an object read as a value when its type has no default
// member to give one (issue #183). Each raising sample was measured through
// pyVBAharness on 2026-09-29 in Excel 16.0 (build 20326); each quiet one runs
// there, except the class with a default member, whose attribute the harness
// cannot set.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostics } from '../helpers/diagnostics';
import { analyzeProjectModule } from './helpers';

const CODE = 'object-default-value';

const NO_DEFAULT = 'Option Explicit\nPublic X As Long\n';
const WITH_DEFAULT =
	'Option Explicit\nPrivate m As Long\n' +
	'Public Property Get X() As Long\nAttribute X.VB_UserMemId = 0\n    X = m\nEnd Property\n' +
	'Public Property Let X(ByVal value As Long)\n    m = value\nEnd Property\n';

function withClass(classSource: string, ...lines: string[]): ReturnType<typeof analyzeModule> {
	const src = `Option Explicit\nSub Main()\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Sub\n`;
	return analyzeProjectModule(src, [{ moduleName: 'Class1', type: 'class', source: classSource }], 'Module1');
}

describe('object-default-value (issue #183)', () => {
	it('flags a class with no default member read as a value, with 438', () => {
		const lines = [
			'Dim c As New Class1, s As String, v As Variant, n As Long',
			's = c',
			'v = c',
			'Debug.Print "a"; c',
			's = c & "x"',
			'n = c + 1',
			'If c = 1 Then Debug.Print 1',
		];
		const src = `Option Explicit\nSub Main()\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Sub\n`;
		const hits = byCode(withClass(NO_DEFAULT, ...lines), CODE);
		expectDiagnostics(src, hits, CODE, Array.from({ length: 6 }, () => ({ span: 'c', message: ["'c' is a Class1, which has no default member", "error '438'"] })));
		// An As New variable is never Nothing, so the message names no 91.
		expect(hits.every((hit) => !hit.message.includes("'91'"))).toBe(true);
	});

	it('names 91 too when the variable may still be Nothing', () => {
		const hits = byCode(withClass(NO_DEFAULT, 'Dim c As Class1, s As String', 'Set c = New Class1', 's = c'), CODE);
		expect(hits).toHaveLength(1);
		expect(hits[0].message).toContain("error '438'");
		expect(hits[0].message).toContain("'91' while it is Nothing");
	});

	it('flags a Collection read whole, with 450, and leaves its operators to collection-operand', () => {
		const src = 'Option Explicit\nSub Main()\n    Dim c As New Collection, v As Variant, n As Long\n    Debug.Print c\n    v = c\n    n = c + 1\nEnd Sub\n';
		const diags = analyzeModule(src);
		expectDiagnostics(src, diags, CODE, [
			{ span: 'c', message: ['default member Item needs an index', "error '450'"] },
			{ span: 'c' },
		]);
		expect(byCode(diags, 'collection-operand')).toHaveLength(1);
	});

	it('flags a Worksheet read as a value', () => {
		const src = 'Option Explicit\nSub Main()\n    Dim ws As Worksheet, s As String\n    Set ws = ActiveSheet\n    s = ws\nEnd Sub\n';
		expectDiagnostics(src, analyzeModule(src), CODE, [{ span: 'ws', message: ["'ws' is a Worksheet", "error '438'"] }]);
	});

	it('stays quiet where no value is read, or the type has one to give', () => {
		const quiet = withClass(
			NO_DEFAULT,
			'Dim c As New Class1, o As Object, r As Range, v As Variant',
			'Set o = c',
			'Take c',
			'If c Is Nothing Then Debug.Print 0',
			'Debug.Print c.X',
			'Set r = Range("A1")',
			'v = r',
		);
		expect(byCode(quiet, CODE)).toHaveLength(0);
		const withDefault = withClass(WITH_DEFAULT, 'Dim c As New Class1, v As Variant', 'v = c', 'Debug.Print c');
		expect(byCode(withDefault, CODE)).toHaveLength(0);
	});

	it('does not read a Function\'s own name, which inside it is a recursive call', () => {
		const src = 'Option Explicit\nFunction Make() As Collection\n    Dim s As String\n    s = Make\nEnd Function\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});
});
