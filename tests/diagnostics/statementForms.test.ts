// Diagnostics tests: statement compile errors from issue #125. Each refused
// sample was measured in Excel 16.0 (build 20326, 2026-09-25) with the VBE's
// message; each accepted one compiles there.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic, expectDiagnostics } from '../helpers/diagnostics';

function wrap(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

describe('next-variable-mismatch - Next i, j (issue #125)', () => {
	it('pairs each name of a Next list with its loop, innermost first', () => {
		const src = wrap('Dim i As Long, j As Long', 'For i = 1 To 2', '    For j = 1 To 2', 'Next i, j', 'Main = 1');
		expectDiagnostics(src, analyzeModule(src), 'next-variable-mismatch', [{ span: 'j' }, { span: 'i' }]);
		const right = wrap('Dim i As Long, j As Long', 'For i = 1 To 2', '    For j = 1 To 2', 'Next j, i', 'Main = 1');
		expect(byCode(analyzeModule(right), 'next-variable-mismatch')).toHaveLength(0);
		expect(byCode(analyzeModule(right), 'block-missing-closer')).toHaveLength(0);
	});
});

describe('for-each-control-variable-type - over an array (issue #125)', () => {
	it('demands Variant when the source is an array', () => {
		for (const [decl, span] of [['Dim arr(1) As Collection, o As Object', 'o'], ['Dim arr(1) As Long, s As String', 's']]) {
			const src = wrap(decl, `For Each ${span} In arr`, 'Next', 'Main = 1');
			expectDiagnostic(src, analyzeModule(src), 'for-each-control-variable-type', { span, message: 'must be Variant when the source is an array' });
		}
		const quiet = wrap('Dim arr(1) As Long, v As Variant, c As New Collection, o As Object', 'For Each v In arr', 'Next', 'For Each o In c', 'Next', 'Main = 1');
		expect(byCode(analyzeModule(quiet), 'for-each-control-variable-type')).toHaveLength(0);
	});
});

describe('collection-operand and sub-used-as-value (issue #125)', () => {
	it('flags a Collection beside an operator', () => {
		const src = wrap('Dim c As New Collection, x As Variant', 'x = c + 1', 'If c = 1 Then x = 2', 'Main = x');
		expectDiagnostics(src, analyzeModule(src), 'collection-operand', [{ span: 'c', message: "'+'" }, { span: 'c', message: "'='" }]);
		const quiet = wrap('Dim c As Collection, d As New Collection', 'Set c = New Collection', 'Set d = c', 'c.Add 1', 'Main = c(1) + d.Count', 'If c Is d Then Main = 2');
		expect(byCode(analyzeModule(quiet), 'collection-operand')).toHaveLength(0);
	});

	it('flags a Sub read as a value and allows a Function', () => {
		const src = 'Option Explicit\nPrivate Sub Foo()\nEnd Sub\nPrivate Function Bar() As Long\nEnd Function\nFunction Main() As Variant\n    Dim x As Variant\n    x = Foo\n    x = Bar\n    Foo\n    Main = x\nEnd Function\n';
		expectDiagnostics(src, analyzeModule(src), 'sub-used-as-value', [{ span: 'Foo', message: 'Expected Function or variable' }]);
	});
});

describe('Set with a literal, Rem after Then and literal forms (issue #125)', () => {
	it('flags Set v = 5 on a Variant', () => {
		const src = wrap('Dim v As Variant', 'Set v = 5', 'Set v = "x"', 'Set v = Nothing', 'Main = 1');
		expectDiagnostics(src, analyzeModule(src), 'set-requires-object', [{ span: '5', message: 'Object required' }, { span: '"x"' }]);
	});

	it('flags Rem after Then', () => {
		const src = wrap('Dim x As Boolean', 'If x Then Rem note', 'Main = 1');
		expectDiagnostic(src, analyzeModule(src), 'rem-after-then', { span: 'Rem', message: 'Syntax error' });
		const quiet = wrap('Dim x As Boolean', "If x Then Main = 1 ' note", 'Rem a comment', 'Main = 1');
		expect(byCode(analyzeModule(quiet), 'rem-after-then')).toHaveLength(0);
	});

	it('flags 1.5% and 1E400', () => {
		const pct = wrap('Dim x As Double', 'x = 1.5%', 'Main = x');
		expectDiagnostic(pct, analyzeModule(pct), 'suffixed-literal-overflow', { span: '1.5%', message: 'fractional' });
		const big = wrap('Dim x As Double', 'x = 1E400', 'Main = x');
		expectDiagnostic(big, analyzeModule(big), 'float-literal-overflow', { span: '1E400', message: 'Double range' });
		const quiet = wrap('Dim x As Double', 'x = 1.5', 'x = 1E300', 'x = 15%', 'Main = x');
		expect(byCode(analyzeModule(quiet), 'float-literal-overflow')).toHaveLength(0);
		expect(byCode(analyzeModule(quiet), 'suffixed-literal-overflow')).toHaveLength(0);
	});

	it('names a module called as a procedure', () => {
		const src = wrap('Foo', 'Main = 1');
		const diags = analyzeModule(src, {
			moduleName: 'Module1',
			knownProcedures: new Set(['bar', 'main']),
			projectClassMembers: [{ name: 'Foo', kind: 'standardModule', moduleName: 'Foo', members: [{ name: 'Bar', kind: 'method', moduleName: 'Foo' }] }],
		});
		expectDiagnostic(src, diags, 'unknown-call', { span: 'Foo', message: 'not module' });
	});
});

describe('one-line If tails, indexed Set targets and AddressOf (issue #140)', () => {
	// Measured in Excel 16.0 (build 20326, 2026-09-26): the module compiles
	// and Main returns 577.
	it('stays quiet on the lazy-initialization idiom and Set into an array element', () => {
		const src =
			'Option Explicit\nPrivate mItems As Collection\n\n' +
			'Public Sub TimerProc()\nEnd Sub\n\n' +
			'Private Function Pointer(ByVal p As LongPtr) As LongPtr\n    Pointer = p\nEnd Function\n\n' +
			'Function Main() As Long\n' +
			'    Dim cols(1 To 2) As Collection, c As Collection, d As Collection, p As LongPtr\n' +
			'    If mItems Is Nothing Then Set mItems = New Collection\n' +
			'    Set d = New Collection\n    d.Add 5\n' +
			'    If d.Count > 5 Then Set c = New Collection Else Set c = d\n' +
			'    Set cols(1) = c\n' +
			'    p = Pointer(AddressOf TimerProc)\n' +
			'    Main = mItems.Count + cols(1).Count\n' +
			'End Function\n';
		const diags = analyzeModule(src);
		expect(byCode(diags, 'collection-operand')).toHaveLength(0);
		expect(byCode(diags, 'sub-used-as-value')).toHaveLength(0);
	});

	it('still judges the condition of a one-line If', () => {
		const src = wrap('Dim c As New Collection', 'If c = 1 Then Main = 2');
		expectDiagnostic(src, analyzeModule(src), 'collection-operand', { span: 'c' });
	});
});
