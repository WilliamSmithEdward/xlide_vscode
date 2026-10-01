// Diagnostics tests: UBound and LBound of something that is no array
// variable or call, which does not compile, and Join over an array whose
// elements are neither Strings nor Variants, or of two dimensions, which
// raises 5. Measured in Excel 16.0 (build 20326, 2026-09-30); each quiet
// neighbour runs clean there.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

function wrap(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

describe('UBound and LBound of no array variable, member or call', () => {
	it.each([
		['UBound(5)', '5'],
		['UBound("abc")', '"abc"'],
		['UBound(1.5)', '1.5'],
		['UBound(#1/1/2000#)', '#1/1/2000#'],
		['UBound(Null)', 'Null'],
		['UBound(True)', 'True'],
		['UBound(Nothing)', 'Nothing'],
		['UBound(Empty)', 'Empty'],
		['UBound(-5)', '-5'],
		['UBound(5 + 1)', '5 + 1'],
		['UBound(5, 1)', '5'],
		['UBound(Len("x"))', 'Len("x")'],
		['UBound(New Collection)', 'New Collection'],
		['LBound(5)', '5'],
	])('reports %s as a Syntax error', (call, span) => {
		const src = wrap(`Main = ${call}`);
		expectDiagnostic(src, byCode(analyzeModule(src), 'malformed-statement'), 'malformed-statement', { span, message: 'Syntax error' });
	});

	it('stays quiet on an array, a call that returns one, VBA.UBound and an array in parentheses', () => {
		for (const lines of [
			['Dim a(3) As Long', 'Main = UBound(a)'],
			['Main = UBound(Array(1, 2))'],
			['Main = UBound(Split("a b"))'],
			['Main = VBA.UBound(5)'],
		]) {
			expect(byCode(analyzeModule(wrap(...lines)), 'malformed-statement'), lines.join('; ')).toHaveLength(0);
		}
		// `UBound((a))` keeps its own message.
		const paren = wrap('Main = UBound((5))');
		const hits = byCode(analyzeModule(paren), 'malformed-statement');
		expect(hits).toHaveLength(1);
		expect(hits[0].message).toContain('not in parentheses');
	});
});

describe('Join over an array it cannot join', () => {
	it.each(['Long', 'Integer', 'Double', 'Date', 'Boolean', 'Byte', 'Currency', 'Object'])('reports a fixed array of %s', (type) => {
		const src = wrap(`Dim a(1) As ${type}`, 'Main = Join(a, "-")');
		expectDiagnostic(src, analyzeModule(src), 'runtime-argument-value', { span: 'a', message: [`an array of ${type}`, "'5'"] });
	});

	it('reports an array of two dimensions, even of Strings', () => {
		for (const type of ['Long', 'String']) {
			const src = wrap(`Dim a(1, 1) As ${type}`, 'Main = Join(a, "-")');
			expectDiagnostic(src, analyzeModule(src), 'runtime-argument-value', { span: 'a', message: ['has 2', "'5'"] });
		}
	});

	it('stays quiet on Strings, Variants and a dynamic array', () => {
		for (const lines of [
			['Dim a(1) As String', 'Main = Join(a, "-")'],
			['Dim a(1) As Variant', 'Main = Join(a, "-")'],
			['Dim a(1)', 'Main = Join(a, "-")'],
			// Unallocated, it joins to "", so a dynamic array is left alone.
			['Dim a() As Long', 'Main = Join(a, "-")'],
			['Main = Join(Array(1, 2), "-")'],
		]) {
			expect(byCode(analyzeModule(wrap(...lines)), 'runtime-argument-value'), lines.join('; ')).toHaveLength(0);
		}
		// A dynamic array before it leaves the rest of the procedure checked.
		const after = wrap('Dim a() As Long, b(1) As Long', 'Main = Join(a, "-")', 'Main = Join(b, "-")');
		expectDiagnostic(after, analyzeModule(after), 'runtime-argument-value', { span: 'b', message: 'an array of Long' });
	});

	it('reads a Static array the same way', () => {
		const src = wrap('Static a(1) As Long', 'Main = Join(a, "-")');
		expectDiagnostic(src, analyzeModule(src), 'runtime-argument-value', { span: 'a', message: 'an array of Long' });
	});
});
