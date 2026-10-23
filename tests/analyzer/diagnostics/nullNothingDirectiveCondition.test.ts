// A Null or Nothing compiler-directive condition (issue #1351). Measured in
// 64-bit Excel 16.0 (2026-10-13), compile mode:
//
//  - A condition whose value is Null is "Invalid use of Null" inside an arm
//    the VBE skips too, and as an #ElseIf after a taken arm: #If Null, 1 =
//    Null, Not Null, Null Or False, and #If N after #Const N = Null. Null
//    that three-valued logic decides compiles (Null Or True, Null And False),
//    and so does #Const N = Null that nothing reads.
//  - The literal Nothing as a condition, alone, in brackets or under Not, is
//    "Invalid use of object", compiled and in a skipped arm, and as an
//    #ElseIf. Nothing Is Nothing compiles, and so does #If O after
//    #Const O = Nothing, compiled and skipped.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../../src/analyzer';
import { byCode, expectDiagnostic } from '../../helpers/diagnostics';

const OPTIONS = { conditionalCompilation: { projectConstants: {} } };
const fn = (...lines: string[]) => ['Function Main() As Variant', ...lines, '    Main = 1', 'End Function', ''].join('\n');
const skipped = (condition: string, before = '') => before + fn('#If False Then', `#If ${condition} Then`, '#End If', '#End If');
const nulls = (src: string) => byCode(analyzeModule(src, OPTIONS), 'null-directive-condition');
const nothings = (src: string) => byCode(analyzeModule(src, OPTIONS), 'nothing-directive-condition');

describe('null-directive-condition wherever the line is', () => {
	it.each([
		['#If Null', skipped('Null')],
		['#If 1 = Null', skipped('1 = Null')],
		['#If Not Null', skipped('Not Null')],
		['#If Null Or False', skipped('Null Or False')],
		['#If Null And True', skipped('Null And True')],
		['#If "abc" = Null', skipped('"abc" = Null')],
		['#If N', skipped('N', '#Const N = Null\n')],
		['#If N = "abc"', skipped('N = "abc"', '#Const N = Null\n')],
	])('flags %s Then inside a skipped arm', (directive, src) => {
		expectDiagnostic(src, analyzeModule(src, OPTIONS), 'null-directive-condition', {
			span: `${directive} Then`, message: 'Invalid use of Null',
		});
	});

	it('flags an #ElseIf Null after a taken arm and inside a skipped arm', () => {
		const taken = fn('#If True Then', '#ElseIf Null Then', '#End If');
		expectDiagnostic(taken, analyzeModule(taken, OPTIONS), 'null-directive-condition', { span: '#ElseIf Null Then' });
		expect(nulls(fn('#If False Then', '#If False Then', '#ElseIf Null Then', '#End If', '#End If'))).toHaveLength(1);
	});

	it('stays quiet where Excel compiles', () => {
		expect(nulls(skipped('Null Or True'))).toHaveLength(0);
		expect(nulls(skipped('Null And False'))).toHaveLength(0);
		expect(nulls(`#Const N = Null\n${fn()}`)).toHaveLength(0);
	});

	it('stays quiet for a skipped non-Null condition and reports one reading a Null constant', () => {
		expect(nulls(skipped('Missing + 1'))).toHaveLength(0);
		// A constant holding Null makes every later condition a candidate.
		expect(nulls(skipped('N + 1', '#Const N = Null\n'))).toHaveLength(1);
	});
});

describe('nothing-directive-condition', () => {
	it.each([
		['#If Nothing Then', fn('#If Nothing Then', '#End If')],
		['#If (Nothing) Then', fn('#If (Nothing) Then', '#End If')],
		['#If Not Nothing Then', fn('#If Not Nothing Then', '#End If')],
		['#If Nothing Then', skipped('Nothing')],
		['#If Not Nothing Then', skipped('Not Nothing')],
		['#ElseIf Nothing Then', fn('#If False Then', '#ElseIf Nothing Then', '#End If')],
	])('flags %s', (directive, src) => {
		expectDiagnostic(src, analyzeModule(src, OPTIONS), 'nothing-directive-condition', {
			span: directive, severity: 'error', message: 'Invalid use of object',
		});
	});

	it('stays quiet for Nothing Is Nothing and for a constant set to Nothing', () => {
		expect(nothings(skipped('Nothing Is Nothing'))).toHaveLength(0);
		expect(nothings(`#Const O = Nothing\n${fn('#If O Then', '#End If')}`)).toHaveLength(0);
		expect(nothings(skipped('O', '#Const O = Nothing\n'))).toHaveLength(0);
		expect(nothings(`#Const O = Nothing\n${fn()}`)).toHaveLength(0);
	});

	it('leaves a line it cannot read to the syntax', () => {
		for (const condition of ['Not Nothing Or (Nothing Is Nothing', 'Not Nothing Or', `${'('.repeat(300)}Not Nothing${')'.repeat(300)}`]) {
			expect(nothings(fn(`#If ${condition} Then`, '#End If')), condition.slice(0, 30)).toHaveLength(0);
		}
	});

	it('still flags a literal Nothing beside a constant set to Nothing', () => {
		const src = `#Const O = Nothing\n${fn('#If Not Nothing Or (O Is Nothing) Then', '#End If')}`;
		expect(nothings(src)).toHaveLength(1);
	});
});
