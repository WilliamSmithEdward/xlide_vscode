// Diagnostics tests: parentheses the VBE reads otherwise, or refuses (issue
// #236). Measured through pyVBAharness in Excel 16.0 on 2026-09-30 with a
// full compile; each message below is Excel's.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode } from '../helpers/diagnostics';

const HEAD = 'Dim c As New Collection, d As Collection, n As Long, s As String, a(1) As Long, v As Variant';
const TAIL = 'Private Sub Take(ByVal x As Variant)\nEnd Sub\nPrivate Function G(Optional ByVal a As Variant, Optional ByVal b As Variant) As Variant\n    G = 1\nEnd Function\n';

function main(line: string): string {
	return `Option Explicit\nFunction Main() As Variant\n    ${HEAD}\n    ${line}\n    Main = 1\nEnd Function\n${TAIL}`;
}

/** The one finding with `code`, whose message ends in the VBE's words, and what it marks. */
function marked(line: string, code: string, vbe: string): string {
	const src = main(line);
	const hits = byCode(analyzeModule(src), code);
	expect(hits, line).toHaveLength(1);
	expect(hits[0].message, line).toContain(`This is a VBE compile error: ${vbe}.`);
	return src.slice(hits[0].span.start, hits[0].span.end);
}

describe('an object in parentheses is its default member (issue #236)', () => {
	it.each([
		['Set d = (New Collection)', '(New Collection)'],
		['Set d = (c)', '(c)'],
		['c.Add (c)', '(c)'],
		['With (New Collection)\n    End With', '(New Collection)'],
		['Call Take((c))', '(c)'],
		['Take (c)', '(c)'],
	])('%s', (line, text) => {
		expect(marked(line, 'collection-operand', 'Argument not optional')).toBe(text);
	});

	it.each(['Set d = c', 'c.Add c', 'v = (n)', 'Take c', 'Call Take(c)', 'With c\n    End With'])('leaves %s alone', (line) => {
		expect(byCode(analyzeModule(main(line)), 'collection-operand')).toEqual([]);
	});
});

describe('malformed-statement: parentheses where VBA takes none (issue #236)', () => {
	it.each([
		['v = (Range("A1")).Address', '.', 'Syntax error'],
		['v = (Application).Name', '.', 'Syntax error'],
		['If (Application).Visible Then v = 1', '.', 'Syntax error'],
		['Debug.Print (Range("A1")).Address', '.', 'Invalid or unqualified reference'],
		['v = ()', '(', 'Syntax error'],
		['v = UBound((a))', '(', 'Syntax error'],
		['v = LBound((a), 1)', '(', 'Syntax error'],
		['v = UBound((a) + 0)', '(', 'Syntax error'],
		['v = Not (Application).Visible', '.', 'Syntax error'],
		['v = G((b:=1), a:=2)', ':=', 'Syntax error'],
		['Mid((s), 1, 1) = "x"', '(', 'Syntax error'],
		['(n) = 1', '(', 'Syntax error'],
		['(Range("A1")).Value = 1', '(', 'Syntax error'],
	])('%s', (line, text, vbe) => {
		expect(marked(line, 'malformed-statement', vbe)).toBe(text);
	});

	it.each([
		'v = Range("A1").Address',
		'v = Array(1, 2)(0)',
		'v = CStr(1)',
		'v = UBound(a)',
		'v = G(b:=(1), a:=2)',
		'v = (G(b:=1))',
		'Mid(s, 1, 1) = "x"',
		'Debug.Print Mid$(String:="abc", Start:=1)',
		'Debug.Print (1), (2)',
		'If (1 = 1) Then v = 1',
		'v = -(1 + 2) * Not (n)',
		'Take (n)',
		'v = ([A1])',
	])('leaves %s alone', (line) => {
		expect(byCode(analyzeModule(main(line)), 'malformed-statement')).toEqual([]);
	});
});

describe('As followed by parentheses (issue #236)', () => {
	it('is Syntax error', () => {
		const hits = byCode(analyzeModule(main('Dim x As (Long)')), 'unexpected-declaration-token');
		expect(hits.map((d) => d.message)).toEqual([expect.stringContaining('This is a VBE compile error: Syntax error.')]);
	});
});
