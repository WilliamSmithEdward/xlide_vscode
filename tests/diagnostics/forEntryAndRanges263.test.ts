// Diagnostics tests: For converts its start, limit and step on entry; a
// counter's range in the value rules; a collection changed while counted;
// and a local stepped past its type (issue #263). Every case was measured in
// Excel 16.0 (build 20326, 2026-10-01).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

function source(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

function errors(src: string): string[] {
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

const loop = (decl: string, header: string, ...body: string[]): string => source(decl, 'Dim n As Long', header, ...body.map((line) => `    ${line}`), 'Next', 'Main = n');

describe('For converts its start, limit and step as it starts', () => {
	it.each([
		['Dim b As Byte', 'For b = 5 To 3 Step -1', '-1', 'converts its step -1 to Byte'],
		['Dim b As Byte', 'For b = 3 To 5 Step -1', '-1', 'converts its step -1 to Byte'],
		['Dim b As Byte', 'For b = -1 To 3', '-1', 'converts its start -1 to Byte'],
		['Dim b As Byte', 'For b = 0 To -1', '-1', 'converts its limit -1 to Byte'],
		['Dim i As Integer', 'For i = 40000 To 1', '40000', 'converts its start 40000 to Integer'],
		['Dim i As Integer', 'For i = 1 To 5 Step 40000', '40000', 'converts its step 40000 to Integer'],
		['Dim i As Long', 'For i = 1 To 3000000000#', '3000000000#', 'converts its limit 3000000000 to Long'],
	])('reports %s: %s', (decl, header, span, message) => {
		const src = loop(decl, header, 'n = n + 1');
		expectDiagnostic(src, byCode(analyzeModule(src), 'for-counter-overflow'), 'for-counter-overflow', { span, message });
	});

	it('reports a limit past the type whatever Exit For the body holds', () => {
		for (const [decl, header] of [['Dim b As Byte', 'For b = 1 To 300'], ['Dim i As Integer', 'For i = 1 To 40000']]) {
			const src = loop(decl, header, 'n = n + 1', 'If n = 5 Then Exit For');
			expectDiagnostic(src, byCode(analyzeModule(src), 'for-counter-overflow'), 'for-counter-overflow', { message: 'converts its limit' });
		}
		const now = loop('Dim i As Integer', 'For i = 1 To 40000', 'Exit For');
		expect(byCode(analyzeModule(now), 'for-counter-overflow')).toHaveLength(1);
	});

	it('reads a step or a limit through a local', () => {
		const step = source('Dim b As Byte, n As Long, s As Long', 's = -1', 'For b = 5 To 3 Step s', '    n = n + 1', 'Next', 'Main = n');
		expectDiagnostic(step, byCode(analyzeModule(step), 'for-counter-overflow'), 'for-counter-overflow', { span: 's', message: 'converts its step -1 to Byte' });
		const top = source('Dim i As Integer, n As Long, top As Long', 'top = 40000', 'For i = 1 To top', '    n = n + 1', '    If n = 10 Then Exit For', 'Next', 'Main = n');
		expectDiagnostic(top, byCode(analyzeModule(top), 'for-counter-overflow'), 'for-counter-overflow', { span: 'top', message: 'converts its limit 40000 to Integer' });
	});

	it('stays quiet where every part fits', () => {
		for (const [decl, header, ...body] of [
			['Dim b As Byte', 'For b = 1 To 3 Step 1', 'n = n + 1'],
			['Dim b As Byte', 'For b = 1 To 5 Step 2', 'n = n + 1'],
			['Dim b As Byte', 'For b = 5 To 3', 'n = n + 1'],
			['Dim b As Byte', 'For b = 5 To 3 Step -0.4', 'n = n + 1', 'If n > 5 Then Exit For'],
			['Dim i As Integer', 'For i = 1 To 32767', 'n = n + 1', 'If n = 10 Then Exit For'],
			['Dim i As Integer', 'For i = 3 To 1 Step -1', 'n = n + 1'],
			['Dim v As Variant', 'For v = 1 To 40000', 'n = n + 1', 'If n = 10 Then Exit For'],
		]) {
			expect(errors(loop(decl, header, ...body)), header).toEqual([]);
		}
	});
});

describe('the counter in the value rules', () => {
	it('divides by a counter that starts at 0', () => {
		const src = source('Dim i As Long, t As Double', 'For i = 0 To 3', '    t = t + 1 / i', 'Next', 'Main = t');
		expectDiagnostic(src, byCode(analyzeModule(src), 'division-by-zero'), 'division-by-zero', { span: 'i', message: "On the first pass of the For loop, where 'i' is 0: " });
	});

	it('converts a counter past Integer', () => {
		const src = source('Dim i As Long, t As Long', 'For i = 32760 To 32770', '    t = CInt(i)', 'Next', 'Main = t');
		expectDiagnostic(src, byCode(analyzeModule(src), 'arithmetic-overflow'), 'arithmetic-overflow', { message: "On the last pass of the For loop, where 'i' is 32770: CInt(32770) does not fit Integer" });
	});

	it.each([
		['a(i) = a(i + 1)', 'i + 1', "Counter 'i' reaches 3 on its last pass, so i + 1 is 4, which for array 'a' is above the upper bound 3"],
		['a(i) = a(i - 1)', 'i - 1', "Counter 'i' is 0 on its first pass, so i - 1 is -1, which for array 'a' is below the lower bound 0"],
	])('indexes a counter a step off: %s', (line, span, message) => {
		const src = source('Dim a(3) As Long, i As Long', 'For i = 0 To 3', `    ${line}`, 'Next', 'Main = a(0)');
		expectDiagnostic(src, byCode(analyzeModule(src), 'array-subscript-out-of-bounds'), 'array-subscript-out-of-bounds', { span, message });
	});

	it.each([
		['For i = 0 To 3', 'i is 4 here'],
		['For i = 0 To 3 Step 2', 'i is 4 here'],
		['For i = 3 To 0 Step -1', 'i is -1 here'],
	])('knows the counter after %s runs out', (header, message) => {
		const src = source('Dim a(3) As Long, i As Long', header, 'Next', 'Main = a(i)');
		expectDiagnostic(src, byCode(analyzeModule(src), 'array-subscript-out-of-bounds'), 'array-subscript-out-of-bounds', { span: 'i', message: `Subscript ${message}` });
	});

	it('stays quiet in range, guarded, after an Exit For, or when no pass runs', () => {
		for (const lines of [
			['Dim i As Long, t As Double', 'For i = 1 To 3', '    t = t + 1 / i', 'Next', 'Main = t'],
			['Dim i As Long, t As Double', 'For i = 0 To 3', '    If i > 0 Then t = t + 1 / i', 'Next', 'Main = t'],
			['Dim i As Long, t As Long', 'For i = 32760 To 32767', '    t = CInt(i)', 'Next', 'Main = t'],
			['Dim a(3) As Long, i As Long', 'For i = 0 To 2', '    a(i) = a(i + 1)', 'Next', 'Main = a(0)'],
			['Dim a(3) As Long, i As Long', 'For i = 1 To 3', '    a(i) = a(i - 1)', 'Next', 'Main = a(0)'],
			['Dim a(3) As Long, i As Long', 'For i = 0 To 3', 'Next', 'Main = a(i - 1)'],
			['Dim a(3) As Long, i As Long', 'For i = 0 To 3', '    If i = 2 Then Exit For', 'Next', 'Main = a(i)'],
			['Dim a(3) As Long, i As Long', 'For i = 2 To 1', 'Next', 'Main = a(i)'],
			['Dim a(3) As Long, i As Long', 'For i = 0 To 3', '    i = i + 0', 'Next', 'Main = 1'],
		]) {
			expect(errors(source(...lines)), lines.join(': ')).toEqual([]);
		}
	});
});

describe('a collection changed while counted', () => {
	const three = ['Dim c As New Collection, i As Long, t As Long', 'c.Add 1: c.Add 2: c.Add 3'];

	it.each([
		[[...three, 'For i = 1 To c.Count', '    c.Remove i', 'Next'], "On the pass of the For loop where 'i' is 3, 'c' holds 1 element, indexed 1 to 1; 3 is outside that. This will raise Run-time error '9'"],
		[['Dim c As New Collection, i As Long', 'c.Add 1: c.Add 2', 'For i = 1 To 3', '    c.Remove 1', 'Next'], "On the pass of the For loop where 'i' is 3, 'c' holds nothing, so no index reaches an element. This will raise Run-time error '5'"],
		[[...three, 'For i = 1 To c.Count', '    t = t + c(i)', '    c.Remove 1', 'Next'], "On the pass of the For loop where 'i' is 3, 'c' holds 1 element"],
	])('reports %j', (lines, message) => {
		const src = source(...lines, 'Main = c.Count');
		expectDiagnostic(src, byCode(analyzeModule(src), 'collection-index-out-of-range'), 'collection-index-out-of-range', { message });
	});

	it('stays quiet removing from the end, removing the first, or with a body it cannot follow', () => {
		for (const lines of [
			[...three, 'For i = c.Count To 1 Step -1', '    c.Remove i', 'Next'],
			[...three, 'For i = 1 To c.Count', '    c.Remove 1', 'Next'],
			[...three, 'For i = 1 To c.Count', '    If i < 3 Then c.Remove i', 'Next'],
			[...three, 'For i = 1 To c.Count', '    c.Remove i', '    c.Add 9', 'Next'],
			[...three, 'For i = 1 To c.Count', '    Debug.Print i', '    c.Remove i', 'Next'],
		]) {
			expect(errors(source(...lines, 'Main = c.Count')), lines.join(': ')).not.toContain('collection-index-out-of-range');
		}
	});
});

describe('a local stepped past its type', () => {
	it.each([
		[['Dim t As Integer, i As Long', 'For i = 1 To 300', '    t = t + i', 'Next', 'Main = t'], 't = t + i', "On the pass of the For loop where 'i' is 256, 't = t + i' makes 't' 32896, which does not fit an Integer"],
		[['Dim p As Long, i As Long', 'p = 1', 'For i = 1 To 20', '    p = p * i', 'Next', 'Main = p'], 'p = p * i', "where 'i' is 13, 'p = p * i' makes 'p' 6227020800, which does not fit a Long"],
		[['Dim i As Integer', 'Do', '    i = i + 1', 'Loop Until i > 32767', 'Main = i'], 'i = i + 1', "the loop ends only when i > 32767, which an Integer never is, so 'i = i + 1' runs until it does not fit"],
		[['Dim b As Byte', 'While b <= 255', '    b = b + 1', 'Wend', 'Main = b'], 'b = b + 1', 'the loop runs while b <= 255, which a Byte always is'],
		[['Dim i As Integer', 'Do While i <= 32767', '    i = i + 1', 'Loop', 'Main = i'], 'i = i + 1', 'the loop runs while i <= 32767'],
		[['Dim b As Byte', 'b = 5', 'Do Until b < 0', '    b = b - 1', 'Loop', 'Main = b'], 'b = b - 1', 'the loop ends only when b < 0, which a Byte never is'],
	])('reports %j', (lines, span, message) => {
		const src = source(...lines);
		expectDiagnostic(src, byCode(analyzeModule(src), 'arithmetic-overflow'), 'arithmetic-overflow', { span, message });
	});

	it('stays quiet when the type holds it, or the loop can end another way', () => {
		for (const lines of [
			['Dim t As Integer, i As Long', 'For i = 1 To 200', '    t = t + i', 'Next', 'Main = t'],
			['Dim p As Long, i As Long', 'p = 1', 'For i = 1 To 12', '    p = p * i', 'Next', 'Main = p'],
			['Dim i As Integer', 'Do', '    i = i + 1', 'Loop Until i >= 32767', 'Main = i'],
			['Dim b As Byte', 'While b < 255', '    b = b + 1', 'Wend', 'Main = b'],
			['Dim i As Long', 'Do While i <= 32767', '    i = i + 1', 'Loop', 'Main = i'],
			['Dim i As Integer', 'Do', '    i = i + 1', '    If i = 100 Then Exit Do', 'Loop Until i > 32767', 'Main = i'],
			['Dim t As Integer, i As Long', 'For i = 1 To 300', '    t = t + i', '    If t > 30000 Then Exit For', 'Next', 'Main = t'],
			['Dim t As Integer, i As Long', 'For i = 1 To 300', '    t = t + i', '    t = 0', 'Next', 'Main = t'],
		]) {
			expect(errors(source(...lines)), lines.join(': ')).toEqual([]);
		}
	});
});
