// Compiler-directive forms the VBE refuses (issue #130). Measured in Excel
// 16.0 (build 20326, 2026-09-26).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

describe('duplicate-const-directive and directive-trailing-statement (issue #130)', () => {
	it('flags a #Const defined twice in one module', () => {
		const src = 'Option Explicit\n#Const FEATURE = 1\n#Const FEATURE = 2\nFunction Main() As Variant\n    Main = 1\nEnd Function\n';
		expectDiagnostic(src, analyzeModule(src), 'duplicate-const-directive', { span: 'FEATURE', message: 'Duplicate definition' });
	});

	it('flags code after the colon on a directive line', () => {
		const src = 'Option Explicit\nFunction Main() As Variant\n#If VBA7 Then: Debug.Print 1\n#End If\n    Main = 1\nEnd Function\n';
		expectDiagnostic(src, analyzeModule(src), 'directive-trailing-statement', { span: 'Debug.Print 1', message: 'whole line' });
	});

	it('stays quiet for one #Const, a directive with a trailing comment and ordinary #If arms', () => {
		const src = "Option Explicit\n#Const FEATURE = 1\nFunction Main() As Variant\n#If FEATURE Then ' the feature\n    Main = 1\n#Else\n    Main = 2\n#End If\nEnd Function\n";
		expect(byCode(analyzeModule(src), 'duplicate-const-directive')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'directive-trailing-statement')).toHaveLength(0);
	});
});

describe('null-directive-condition (issue #208)', () => {
	// Measured in 64-bit Excel 16.0: "Invalid use of Null" for each.
	const wrap = (setup: string, directive: string): string =>
		`Option Explicit\n${setup}Function Main() As Variant\n${directive}\n    Main = 1\n#Else\n    Main = 2\n#End If\nEnd Function\n`;
	const nulls = (src: string) => byCode(analyzeModule(src, { conditionalCompilation: { projectConstants: {} } }), 'null-directive-condition');

	it.each([
		['', '#If Null Then'],
		['', '#If Null = 1 Then'],
		['', '#If Not Null Then'],
		['', '#If Null Like "a" Then'],
		['', '#If "a" Like Null Then'],
		['#Const N = Null\n', '#If N Then'],
	])('flags %s%s', (setup, directive) => {
		const src = wrap(setup, directive);
		expectDiagnostic(src, analyzeModule(src), 'null-directive-condition', { span: directive, message: 'Invalid use of Null' });
	});

	it('flags an #ElseIf the VBE reaches, and not one after a True arm', () => {
		const reached = 'Function Main() As Variant\n#If False Then\n    Main = 1\n#ElseIf Null Then\n    Main = 2\n#End If\nEnd Function\n';
		expect(nulls(reached)).toHaveLength(1);
		expect(nulls(reached.replace('#If False', '#If True'))).toHaveLength(0);
	});

	it('stays quiet where Null is decided, and inside a branch that is not compiled', () => {
		expect(nulls(wrap('', '#If Null Or True Then'))).toHaveLength(0);
		expect(nulls(wrap('', '#If Null And False Then'))).toHaveLength(0);
		expect(nulls(wrap('', '#If Null & "a" = "a" Then'))).toHaveLength(0);
		expect(nulls('#If False Then\n#If Null Then\n#End If\n#End If\n')).toHaveLength(0);
	});
});
