// Diagnostics tests: member and name uses the VBE refuses (issue #369). Every
// case was measured in Excel 16.0 (build 20326, 2026-10-01), compiled with
// the VBE's Debug > Compile.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';
import { analyzeProjectModule, type ProjectTestModule } from './helpers';

function source(decl: string, ...lines: string[]): string {
	return `Option Explicit\n${decl}Function Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

function errors(diagnostics: ReturnType<typeof analyzeModule>): string[] {
	return diagnostics.filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a variable As Sheets or Worksheets used as a value', () => {
	it('is Argument not optional', () => {
		for (const [type, line, span] of [
			['Sheets', 's = o', 'o'],
			['Sheets', 'Main = o & "x"', 'o'],
			['Sheets', 'If o = 7 Then Main = 1', 'o'],
			['Worksheets', 's = o', 'o'],
		] as const) {
			const src = source('', `Dim o As ${type}, s As String`, 'Set o = Worksheets', line);
			expectDiagnostic(src, byCode(analyzeModule(src), 'collection-operand'), 'collection-operand', { span, message: 'Argument not optional' });
		}
		expect(errors(analyzeModule(source('', 'Dim o As Sheets', 'Set o = Worksheets', 'Main = o.Count')))).toEqual([]);
	});
});

describe('a class Sub read as a value', () => {
	const modules = (body: string): ProjectTestModule[] => [{ moduleName: 'Class1', moduleKind: 'class', source: `Option Explicit\n${body}` }];
	it('is Expected Function or variable', () => {
		const src = source('', 'Dim c As New Class1, x As Variant', 'x = c.DoIt()', 'Main = 1');
		const diagnostics = analyzeProjectModule(src, modules('Public Sub DoIt()\nEnd Sub\n'), 'Module1');
		expectDiagnostic(src, byCode(diagnostics, 'sub-used-as-value'), 'sub-used-as-value', { span: 'DoIt', message: 'Expected Function or variable' });
	});

	it('leaves a Function alone', () => {
		const src = source('', 'Dim c As New Class1, x As Variant', 'x = c.DoIt()', 'Main = x');
		const diagnostics = analyzeProjectModule(src, modules('Public Function DoIt() As Long\n    DoIt = 3\nEnd Function\n'), 'Module1');
		expect(byCode(diagnostics, 'sub-used-as-value')).toEqual([]);
	});
});

describe('a module named like a procedure, used bare from another module', () => {
	const foo = (body: string): ProjectTestModule[] => [{ moduleName: 'Foo', source: `Option Explicit\n${body}` }];
	it.each([
		['Foo', 'Public Sub Foo()\nEnd Sub\n'],
		['Call Foo', 'Public Sub Foo()\nEnd Sub\n'],
		['Foo 1, 2', 'Public Sub Foo(a As Long, b As Long)\nEnd Sub\n'],
		['Main = Foo()', 'Public Function Foo() As Long\n    Foo = 1\nEnd Function\n'],
	])('reports %s', (line, body) => {
		const src = source('', line, 'Main = 1');
		const diagnostics = analyzeProjectModule(src, foo(body), 'Module1');
		expectDiagnostic(src, byCode(diagnostics, 'malformed-statement'), 'malformed-statement', { span: 'Foo', message: 'not module' });
	});

	it('leaves Foo.Foo, a call inside Foo, and the caller\'s own names alone', () => {
		const qualified = source('', 'Foo.Foo', 'Main = 1');
		expect(byCode(analyzeProjectModule(qualified, foo('Public Sub Foo()\nEnd Sub\n'), 'Module1'), 'malformed-statement')).toEqual([]);
		const inside = 'Option Explicit\nPublic Sub Foo()\nEnd Sub\nPublic Sub Bar()\n    Foo\nEnd Sub\n';
		expect(byCode(analyzeProjectModule(inside, [{ moduleName: 'Module1', source: source('', 'Main = 1') }], 'Foo'), 'malformed-statement')).toEqual([]);
		for (const decl of ['Private Foo As Long\n', 'Private Sub Foo()\nEnd Sub\n']) {
			const own = source(decl, decl.includes('Sub') ? 'Foo' : 'Foo = 1', 'Main = 1');
			expect(byCode(analyzeProjectModule(own, foo('Public Sub Bar()\nEnd Sub\n'), 'Module1'), 'malformed-statement'), decl).toEqual([]);
		}
	});
});

describe('an Integer field of a Type passed ByRef to a Long', () => {
	it('is ByRef argument type mismatch', () => {
		const decl = (type: string) => `Private Type T1\n    i As ${type}\nEnd Type\nPrivate Sub TakeL(ByRef n As Long)\nEnd Sub\n`;
		const src = source(decl('Integer'), 'Dim t As T1', 'TakeL t.i', 'Main = 1');
		expectDiagnostic(src, byCode(analyzeModule(src), 'byref-argument-type-mismatch'), 'byref-argument-type-mismatch', { span: 't.i' });
		expect(byCode(analyzeModule(source(decl('Long'), 'Dim t As T1', 'TakeL t.i', 'Main = 1')), 'byref-argument-type-mismatch')).toEqual([]);
	});
});

describe('names after VBA.', () => {
	it.each([
		['Main = VBA.Nosuch', 'Nosuch'],
		['Main = VBA.Strings.Nosuch', 'Nosuch'],
		['Main = VBA.Global.Left$("ab", 1)', 'Left'],
		['Main = VBA.VbMsgBoxResult.vbNosuch', 'vbNosuch'],
	])('reports %s', (line, span) => {
		const src = source('', line);
		expectDiagnostic(src, byCode(analyzeModule(src), 'member-not-found'), 'member-not-found', { span, message: 'Method or data member not found' });
	});

	it.each([
		['Main = VBA.Asc$("a")', 'Asc'],
		['Main = VBA.Strings.Asc$("a")', 'Asc'],
		['Main = VBA.Len$("a")', 'Len'],
	])('reports the $ form %s', (line, span) => {
		const src = source('', line);
		expectDiagnostic(src, byCode(analyzeModule(src), 'member-not-found'), 'member-not-found', { span, message: 'Type-declaration character' });
	});

	it('reports VBA.Err.LastDllError as read-only', () => {
		const src = source('', 'VBA.Err.LastDllError = 5', 'Main = 1');
		expectDiagnostic(src, byCode(analyzeModule(src), 'readonly-member-assignment'), 'readonly-member-assignment', { span: 'LastDllError', message: "Can't assign to read-only property" });
	});

	it('leaves the library\'s own names alone', () => {
		for (const line of [
			'Main = VBA.Global.UserForms.Count', 'Main = VBA.Strings.Left$("ab", 1)', 'Main = VBA.Left$("ab", 1)', 'Main = VBA.Left("ab", 1)',
			'Main = VBA.vbYes', 'Main = Len(VBA.vbCrLf)', 'Main = UBound(VBA.Array(1, 2))', 'Main = VBA.Err.Number', 'Main = VBA.Err.Nosuch',
			'Main = VBA.Interaction.IIf(True, 1, 2)', 'Main = VBA.VbMsgBoxResult.vbYes', 'Main = VBA.LCase$("A")', 'Main = VBA.Mid$("abc", 2)',
			'Main = VBA.Strings.Chr(65)', 'Main = VBA.Conversion.CStr(1)', 'Main = Len(VBA.Constants.vbCr)', 'VBA.Err.Description = "x"',
			// A String function takes a $ though the library lists no $ form.
			'Main = VBA.Replace$("ab", "a", "c")', 'Main = VBA.Strings.Replace$("ab", "a", "c")', 'Main = VBA.Join$(Array("a", "b"))', 'Main = VBA.TypeName$(1)',
		]) {
			expect(byCode(analyzeModule(source('', line)), 'member-not-found'), line).toEqual([]);
		}
		const regexp = source('', 'Dim r As New VBA.RegExp', 'r.Pattern = "a"', 'Main = r.Test("ab")');
		expect(errors(analyzeModule(regexp))).toEqual([]);
	});

	it('is not judged where the module names something VBA', () => {
		const src = source('Private Function VBA() As Object\nEnd Function\n', 'Main = VBA.Nosuch');
		expect(byCode(analyzeModule(src), 'member-not-found')).toEqual([]);
	});
});
