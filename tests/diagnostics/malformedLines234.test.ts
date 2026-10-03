// Diagnostics tests: lines the VBE cannot parse, as keyboard mash leaves
// them (issue #234). Measured through pyVBAharness in Excel 16.0 on
// 2026-09-30 with a full compile; each message below is Excel's.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode } from '../helpers/diagnostics';

const MAIN = 'Function Main() As Variant\n    Main = 1\nEnd Function\n';

function inProcedure(line: string): string {
	return `Option Explicit\nFunction Main() As Variant\n    Dim v As Variant\n    ${line}\n    Main = 1\nEnd Function\n`;
}

function errors(src: string): Array<{ code: string; message: string }> {
	return analyzeModule(src)
		.filter((d) => d.severity === 'error')
		.map((d) => ({ code: d.code ?? '', message: d.message }));
}

/** The one finding with `code`, whose message ends in the VBE's words. */
function expectOne(src: string, code: string, vbe: string): void {
	const hits = byCode(analyzeModule(src), code);
	expect(hits, src).toHaveLength(1);
	expect(hits[0].message, src).toContain(`This is a VBE compile error: ${vbe}.`);
}

describe('inactive #If branches are not parsed (issue #234)', () => {
	it.each([
		['an unclosed string', '#If False Then\nasdf "qwer\n#End If\n'],
		['a lone quote', '#If False Then\n"\n#End If\n'],
		['a continuation into a blank line', '#If False Then\nasdf _\n\n#End If\n'],
		['brackets, symbols and a stray #', '#If False Then\n[asdf\n@#$\n#junk\n#End If\n'],
	])('leaves %s alone', (_label, block) => {
		expect(errors(`Option Explicit\n${MAIN}${block}`)).toEqual([]);
	});

	it('still reports an unclosed string in an active branch', () => {
		expect(byCode(analyzeModule(`Option Explicit\n${MAIN}#If True Then\nasdf "qwer _\n#End If\n`), 'unterminated-string')).toHaveLength(1);
	});
});

describe('malformed-statement (issue #234)', () => {
	it.each([
		['#asdf in a procedure', inProcedure('#asdf'), 'Syntax error'],
		['#asdf above the first procedure', `Option Explicit\n#asdf\n${MAIN}`, 'Expected: If or Else or ElseIf or End or EndIf or Const'],
		['#asdf after a procedure', `Option Explicit\n${MAIN}#asdf\n`, 'Syntax error'],
		['an unclosed bracket', inProcedure('[asdf'), 'Syntax error'],
		["an unclosed bracket after Then", inProcedure('If True Then [x'), 'Syntax error'],
		['Sub with no name above the first procedure', `Option Explicit\nSub\n${MAIN}`, 'Expected: identifier'],
		['Function with no name after a procedure', `Option Explicit\n${MAIN}Function\n`, 'Syntax error'],
		['a word after a parameter', `Option Explicit\nPrivate Sub S(a b c)\nEnd Sub\n${MAIN}`, 'Expected: list separator or )'],
		['a keyword after a parameter', `Option Explicit\nPrivate Sub S(x Property)\nEnd Sub\n${MAIN}`, 'Expected: list separator or )'],
		['two words in an Enum', `Option Explicit\nPrivate Enum E\n    eA = 1\n    asdf qwer\nEnd Enum\n${MAIN}`, 'Invalid inside Enum'],
		['a dotted Enum line', `Option Explicit\nPrivate Enum E\n    eA = 1\n    .asdf\nEnd Enum\n${MAIN}`, 'Invalid inside Enum'],
		['an operator after an Enum member', `Option Explicit\nPrivate Enum E\n    eA = 1\n    asdf & qwer\nEnd Enum\n${MAIN}`, 'Expected: expression'],
		['an unclosed bracket in an Enum', `Option Explicit\nPrivate Enum E\n    eA = 1\n    [asdf\nEnd Enum\n${MAIN}`, 'Missing end bracket'],
	])('%s', (_label, src, vbe) => {
		expectOne(src, 'malformed-statement', vbe);
	});

	it('names a dotted Enum line whole', () => {
		const [hit] = byCode(analyzeModule(`Option Explicit\nPrivate Enum E\n    eA = 1\n    .asdf\nEnd Enum\n${MAIN}`), 'malformed-statement');
		expect(hit.message).toMatch(/^'\.asdf' is no Enum member/);
	});

	it.each([
		['#If, #ElseIf, #Else and #End If', `Option Explicit\n#Const A = 1\n#If A Then\n#ElseIf 0 Then\n#Else\n#End If\n${MAIN}`],
		['a bracketed name', inProcedure('v = [A1]')],
		['parameters in every form', `Option Explicit\nPrivate Sub S(ByVal a As Long, c() As Long, Optional b$ = "")\nEnd Sub\nPrivate Sub T(ParamArray d())\nEnd Sub\n${MAIN}`],
		['Enum members with and without values', `Option Explicit\nPrivate Enum E\n    eA = 1\n    eB\n    [e C] = eA + 1\nEnd Enum\n${MAIN}`],
		// "Expected: label or statement or end of statement": invalid-identifier-start's.
		['a number for an Enum member', `Option Explicit\nPrivate Enum E\n    eA = 1\n    123 = 2\nEnd Enum\n${MAIN}`],
	])('leaves %s alone', (_label, src) => {
		expect(byCode(analyzeModule(src), 'malformed-statement')).toEqual([]);
	});
});

describe('reserved-keyword-in-expression (issue #234)', () => {
	it.each([
		['v = Array(1, Const, 3)', 'Const'],
		['Debug.Print RaiseEvent', 'RaiseEvent'],
		['Dim a(Implements) As Long', 'Implements'],
		['v = Loop', 'Loop'],
	])('%s', (line, word) => {
		const src = inProcedure(line);
		expectOne(src, 'reserved-keyword-in-expression', 'Syntax error');
		const [hit] = byCode(analyzeModule(src), 'reserved-keyword-in-expression');
		expect(src.slice(hit.span.start, hit.span.end)).toBe(word);
	});

	it('reports one after Case, and says Expected: expression in a Const', () => {
		expectOne('Option Explicit\nFunction Main() As Variant\n    Select Case 1\n    Case Open\n    End Select\n    Main = 1\nEnd Function\n', 'reserved-keyword-in-expression', 'Syntax error');
		expectOne(`Option Explicit\nPrivate Const K = Dim\n${MAIN}`, 'reserved-keyword-in-expression', 'Expected: expression');
	});

	it.each([
		'Set v = New Collection',
		'Workbooks.Open Filename:="x"',
		'v = Range("A1").End(xlDown).Row',
		'v = ActiveWorkbook.Close',
		'Select Case 1\n    Case Is > 0, 1 To 3\n    Case Else\n    End Select',
		'Dim a(1 To 5) As Long',
		'For v = 1 To 3 Step 1\n    Next',
		'Open "x" For Input As #1',
	])('leaves %s alone', (line) => {
		expect(byCode(analyzeModule(inProcedure(line)), 'reserved-keyword-in-expression')).toEqual([]);
	});
});

describe('declarations that are not name [As type] (issue #234)', () => {
	it.each([
		['Dim asdf qwer in a procedure', inProcedure('Dim asdf qwer'), 'Syntax error'],
		['Dim asdf qwer above the first procedure', `Option Explicit\nDim asdf qwer\n${MAIN}`, 'Expected: end of statement'],
		['Private asdf qwer', `Option Explicit\nPrivate asdf qwer\n${MAIN}`, 'Expected: end of statement'],
		['a Const with two values', `Option Explicit\nPrivate Const K = asdf qwer\n${MAIN}`, 'Expected: end of statement'],
		['As with no type name', `Option Explicit\nPrivate v As 123\n${MAIN}`, 'Expected: New or type name'],
	])('%s', (_label, src, vbe) => {
		expectOne(src, 'unexpected-declaration-token', vbe);
	});

	it.each([
		'Dim a(1 To 3) As Long, b$, c',
		'Const K = 1 + 2',
		'Static s As String * 10',
		// Names the identifier rules report.
		'Dim _name As String',
		'Dim 1value As Long',
		'Dim user-name As String',
	])('leaves %s alone', (line) => {
		expect(byCode(analyzeModule(inProcedure(line)), 'unexpected-declaration-token')).toEqual([]);
	});
});

describe('other lines (issue #234)', () => {
	it('reports Not with nothing to negate', () => {
		expectOne(inProcedure('Not'), 'invalid-expression-syntax', 'Syntax error');
	});

	it('reports a lone colon after a procedure, and not above the first', () => {
		expect(byCode(analyzeModule(`Option Explicit\n${MAIN}:\n`), 'statement-outside-procedure')).toHaveLength(1);
		expect(byCode(analyzeModule(`Option Explicit\n:\n${MAIN}`), 'statement-outside-procedure')).toEqual([]);
	});

	it('names a character outside the Basic Multilingual Plane whole', () => {
		const emoji = String.fromCodePoint(0x1f600);
		const hits = byCode(analyzeModule(inProcedure(`v = 1 ${emoji}`)), 'stray-character');
		expect(hits).toHaveLength(1);
		expect(hits[0].message).toContain(`'${emoji}'`);
		expect(hits[0].span.end - hits[0].span.start).toBe(2);
	});
});
