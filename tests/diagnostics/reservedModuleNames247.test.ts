// Diagnostics tests: a module named a word the VBE refuses, and a call
// qualified by a reserved word (issue #247). Measured in Excel 16.0 on
// 2026-09-30: adding a module named Print, If or String fails with 0x800AC3D4,
// and `Main = Print.Hi()` is a Syntax error with or without such a module.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

const BODY = 'Option Explicit\nPublic Function Hi() As Long\n    Hi = 1\nEnd Function\n';

function wrap(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

function codes(src: string, moduleName?: string): string[] {
	return analyzeModule(src, { moduleName }).map((diag) => diag.code);
}

describe('module names', () => {
	it.each(['Print', 'Date', 'Circle', 'Debug', 'Me', 'Stop', 'If', 'String', 'Nothing', 'LenB'])('reports a module named %s', (name) => {
		const src = `Attribute VB_Name = "${name}"\n${BODY}`;
		expectDiagnostic(src, byCode(analyzeModule(src, { moduleName: name }), 'invalid-declaration-name'), 'invalid-declaration-name', {
			span: `"${name}"`,
			message: `'${name}' cannot name a module`,
		});
	});

	it('marks the first line when the source has no VB_Name attribute', () => {
		const diags = byCode(analyzeModule(BODY, { moduleName: 'print' }), 'invalid-declaration-name');
		expect(diags).toHaveLength(1);
		expect(diags[0].span.start).toBe(0);
	});

	it.each(['Line', 'Width', 'Name', 'Err', 'Mid', 'Time', 'Error', 'Reset', 'Beep', 'Load', 'Unload', 'Access', 'Base', 'Compare', 'Explicit', 'Object', 'Property', 'Step', 'Printer', 'Module1'])('stays quiet on a module named %s', (name) => {
		expect(codes(`Attribute VB_Name = "${name}"\n${BODY}`, name)).not.toContain('invalid-declaration-name');
	});
});

describe('reserved qualifiers', () => {
	it.each(['Circle', 'PSet', 'Scale', 'Print', 'Input', 'Tab', 'Spc', 'Array', 'LBound', 'Date'])('reports %s.Hi()', (word) => {
		const src = wrap(`Main = ${word}.Hi()`);
		expectDiagnostic(src, byCode(analyzeModule(src), 'reserved-keyword-in-expression'), 'reserved-keyword-in-expression', {
			span: word,
			message: 'Syntax error',
		});
	});

	it('reports a qualifier inside a call statement', () => {
		const src = wrap('Call Foo(Print.Hi())');
		expect(codes(src)).toContain('reserved-keyword-in-expression');
	});

	it.each(['Circle', 'PSet', 'Scale', 'Input', 'Tab', 'Spc', 'Array', 'LBound'])('reports %s.Hi opening a statement', (word) => {
		const src = wrap(`${word}.Hi`);
		expectDiagnostic(src, byCode(analyzeModule(src), 'reserved-keyword-in-expression'), 'reserved-keyword-in-expression', {
			span: word,
			message: 'Syntax error',
		});
	});

	it('reports Print.Hi opening a statement with the error the VBE gives', () => {
		const src = wrap('Print.Hi');
		expectDiagnostic(src, byCode(analyzeModule(src), 'reserved-keyword-in-expression'), 'reserved-keyword-in-expression', {
			span: 'Print',
			message: 'Method not valid without suitable object',
		});
	});

	it('stays quiet on Date.Hi opening a statement, which compiles', () => {
		expect(codes(wrap('Date.Hi'))).not.toContain('reserved-keyword-in-expression');
	});

	it('stays quiet on the words as members and on accepted qualifiers', () => {
		for (const statement of [
			'Main = Line.Hi()',
			'Main = Err.Number',
			'Main = Me2.Print',
			'Debug.Print 1',
			'Main = Application.Print',
			'Main = Date',
			'Main = Array(1, 2)',
			'Main = LBound(Array(1))',
		]) {
			expect(codes(wrap(statement)), statement).not.toContain('reserved-keyword-in-expression');
		}
	});
});
