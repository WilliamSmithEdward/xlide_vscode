// Diagnostics tests: a Rem after a statement (issue #231). Measured through
// pyVBAharness in Excel 16.0 on 2026-09-30 with a full compile. In a one-line
// If's Then or Else list it is a comment, even right after Else, and it may
// swallow that If's Else. It is also one after a line number, a label and a
// block Else. Anywhere else it is "Syntax error" in a procedure and
// "Expected: end of statement" at module level or on a procedure's own line.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode } from '../helpers/diagnostics';

const CODE = 'rem-after-statement';

function main(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

function errors(src: string): string[] {
	return analyzeModule(src).filter((d) => d.severity === 'error').map((d) => d.code ?? '');
}

describe('rem-after-statement (issue #231)', () => {
	it.each([
		['after an assignment', ['Main = 1 Rem note']],
		['after Debug.Print', ['Debug.Print 1 Rem note']],
		['after Exit', ['Main = 1', 'Exit Function Rem note']],
		['after GoTo', ['GoTo L Rem note', 'L: Main = 1']],
		['after Next', ['Dim i As Long', 'For i = 1 To 2', 'Next Rem note', 'Main = 1']],
		['after End If', ['If True Then', '    Main = 1', 'End If Rem note']],
		['on a declared name', ['Dim note As Long', 'Main = 1 Rem note']],
		['with nothing after it', ['Main = 1 Rem']],
		['as a value', ['Main = Rem']],
		['after a Dim', ['Dim x As Long Rem note', 'Main = 1']],
		['after a Const', ['Const K = 1 Rem note', 'Main = K']],
		['after Case', ['Select Case 1', 'Case 1 Rem note', '    Main = 1', 'End Select']],
		['after Select Case', ['Select Case 1 Rem note', 'Case 1', '    Main = 1', 'End Select']],
		['after Do', ['Do Rem note', '    Main = 1', '    Exit Do', 'Loop']],
		['after With', ['With New Collection Rem note', '    Main = .Count + 1', 'End With']],
		['after a line number and a statement', ['10 Main = 1 Rem note']],
		['after a statement behind a label', ['L1: Main = 1 Rem note']],
		['on the line after a one-line If', ['If True Then Main = 0', 'Main = 1 Rem note']],
	])('%s: Syntax error', (_label, lines) => {
		const hits = byCode(analyzeModule(main(...lines)), CODE);
		expect(hits).toHaveLength(1);
		expect(hits[0].message).toContain('This is a VBE compile error: Syntax error.');
	});

	it('marks the word Rem', () => {
		const src = main('Main = 1 Rem note');
		const [hit] = byCode(analyzeModule(src), CODE);
		expect(src.slice(hit.span.start, hit.span.end)).toBe('Rem');
	});

	it('reads none of the comment as code', () => {
		expect(errors(main('Main = 1 Rem note'))).toEqual([CODE]);
		expect(errors(main('Dim i As Long', 'For i = 1 To 2', 'Next Rem note', 'Main = 1'))).toEqual([CODE]);
	});

	it.each([
		['a declaration', 'Option Explicit\nPrivate m As Long Rem note\nFunction Main() As Variant\n    Main = m\nEnd Function\n'],
		['a Const', 'Option Explicit\nConst K = 1 Rem note\nFunction Main() As Variant\n    Main = K\nEnd Function\n'],
		['Option Explicit', 'Option Explicit Rem note\nFunction Main() As Variant\n    Main = 1\nEnd Function\n'],
		["a procedure's own line", 'Option Explicit\nFunction Main() As Variant Rem note\n    Main = 1\nEnd Function\n'],
	])('at module level, after %s: Expected: end of statement', (_label, src) => {
		const hits = byCode(analyzeModule(src), CODE);
		expect(hits).toHaveLength(1);
		expect(hits[0].message).toContain('This is a VBE compile error: Expected: end of statement.');
	});

	it('is Syntax error on End Function', () => {
		const hits = byCode(analyzeModule('Option Explicit\nFunction Main() As Variant\n    Main = 1\nEnd Function Rem note\n'), CODE);
		expect(hits.map((hit) => hit.message)).toEqual([expect.stringContaining('Syntax error.')]);
	});

	it.each([
		['a one-line If', ['If True Then Main = 1 Rem note']],
		["a one-line If's Else", ['If False Then Main = 1 Else Main = 2 Rem note']],
		['right after Else', ['If True Then Main = 1 Else Rem note']],
		['a later statement of the list', ['If True Then Main = 1: Main = 2 Rem note']],
		['one that swallows the Else', ['If False Then Main = 1 Rem note Else Main = 2']],
		['nothing after it', ['If True Then Main = 1 Rem']],
		['REM in capitals', ['If True Then Main = 1 REM note']],
		['a continued comment', ['If True Then Main = 1 Rem note _', '    more']],
		['after GoTo in the list', ['If True Then GoTo L Rem note', 'L:  Main = 1']],
		["the Else list's later statement", ['If True Then Main = 1 Else Main = 2: Main = 3 Rem note']],
		['after a colon', ['Main = 1: Rem note']],
		['a list opened by Then and a colon', ['If True Then: Main = 1 Rem note']],
		['after a line number', ['10 Rem note', 'Main = 1']],
		['after a label', ['L1: Rem note', 'Main = 1']],
		['after a block Else', ['If False Then', 'Else Rem note', '    Main = 1', 'End If']],
		['an apostrophe comment naming Rem', ["Main = 1 ' Rem note"]],
	])('leaves a comment in %s', (_label, lines) => {
		expect(errors(main(...lines))).toEqual([]);
	});

	it.each([
		['If', ['If True Then Rem note']],
		['ElseIf', ['If False Then', 'ElseIf True Then Rem note', '    Main = 1', 'End If']],
	])('leaves rem-after-then to judge Rem right after Then: %s', (_label, lines) => {
		const src = main(...lines);
		expect(byCode(analyzeModule(src), 'rem-after-then')).toHaveLength(1);
		expect(byCode(analyzeModule(src), CODE)).toEqual([]);
	});

	it('leaves a line in an inactive #If branch alone', () => {
		expect(byCode(analyzeModule(main('#If False Then', 'Main = 1 Rem note', '#End If', 'Main = 1')), CODE)).toEqual([]);
	});
});
