// Diagnostics tests: lexical rule family.
// Split verbatim from tests/vbaDiagnostics.test.ts (audit #107).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

import { byCode, spanText } from '../helpers/diagnostics';

describe('analyzeModule - unterminated string', () => {
	// Issue #681, measured in Excel 16.0: the VBE closes a string left open at
	// the end of its line, unless the line ends like a continuation.
	it('flags an open string whose line ends with a continuation', () => {
		const src = 'Sub T()\n    MsgBox "hello _\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'unterminated-string');
		expect(hits.length).toBe(1);
		expect(spanText(src, hits[0])).toBe('"hello _');
		expect(hits[0].severity).toBe('error');
		expect(hits[0].message).toContain('Syntax error');
	});

	it.each([
		['  _', 'Main = "abc  _'],
		['a tab', 'Main = "abc\t_'],
		['a blank after the underscore', 'Main = "abc _ '],
		['an open string after a closed one', 'Main = "abc" & "x _'],
		['nothing but the continuation', 'Main = " _'],
	])('flags %s', (_label, line) => {
		expect(byCode(analyzeModule(`Function Main()\n    ${line}\nEnd Function\n`), 'unterminated-string')).toHaveLength(1);
	});

	it.each([
		'Main = "abc',
		'Main = "ab" & "cd',
		'Main = "abc \' not a comment',
		'Dim s As String: s = "x": Main = s & "y',
		'If True Then Main = "abc',
		'Main = "a""b',
		'Main = "abc" & "',
		'Main = "abc_',
		'Main = "abc __',
		'Main = "a _b',
		'Main = "_',
	])('leaves %s alone, which the VBE closes', (line) => {
		expect(byCode(analyzeModule(`Function Main()\n    ${line}\nEnd Function\n`), 'unterminated-string')).toHaveLength(0);
	});

	it('leaves a Const left open alone', () => {
		expect(byCode(analyzeModule('Private Const K As String = "abc\nFunction Main()\n    Main = K\nEnd Function\n'), 'unterminated-string')).toHaveLength(0);
	});

	it('reports a string that swallowed a parenthesis as unbalanced', () => {
		const diagnostics = analyzeModule('Function Main()\n    Main = Len("abc\nEnd Function\n');
		expect(byCode(diagnostics, 'unterminated-string')).toHaveLength(0);
		expect(byCode(diagnostics, 'unbalanced-parens')).toHaveLength(1);
	});

	it('accepts a properly closed string, including doubled-quote escapes', () => {
		const src = 'Sub T()\n    MsgBox "say ""hi"" now"\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'unterminated-string')).toHaveLength(0);
	});

	it('treats a trailing escaped pair without a real close as unterminated', () => {
		const src = 'Sub T()\n    x = "ab"" _\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'unterminated-string')).toHaveLength(1);
	});

	// Corpus RT_003 (excel_vba_realtime_analysis_test_corpus.md): an unterminated
	// string while typing stays a single local diagnostic and clears once closed.
	it('reports an unterminated string while typing, then clears when the quote is closed', () => {
		const opening = 'Sub T()\n    Debug.Print " _\nEnd Sub\n';
		const partial = 'Sub T()\n    Debug.Print "hello _\nEnd Sub\n';
		const closed = 'Sub T()\n    Debug.Print "hello" _\n        & "x"\nEnd Sub\n';
		expect(byCode(analyzeModule(opening), 'unterminated-string')).toHaveLength(1);
		const partialHits = byCode(analyzeModule(partial), 'unterminated-string');
		expect(partialHits).toHaveLength(1);
		expect(spanText(partial, partialHits[0])).toBe('"hello _');
		expect(byCode(analyzeModule(closed), 'unterminated-string')).toHaveLength(0);
	});
});

describe('analyzeModule - invalid line continuation', () => {
	it('flags a continuation underscore followed by a trailing comment', () => {
		const src =
			'Sub T()\n' +
			'    Debug.Print "hello" & _ \' bad trailing comment\n' +
			'        "world"\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-line-continuation');

		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe("_ ' bad trailing comment");
		expect(hits[0].severity).toBe('error');
	});

	it('flags a continuation underscore followed by more code on the same line', () => {
		const src = 'Sub T()\n    total = 1 _ + 2\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-line-continuation');

		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('_ + 2');
	});

	it('flags a likely continuation with no whitespace before the underscore', () => {
		const src =
			'Sub T()\n' +
			'    Debug.Print "hello" &_\n' +
			'        "world"\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-line-continuation');

		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('_');
	});

	it('accepts valid continuations', () => {
		const src =
			'Sub T()\n' +
			'    Debug.Print "hello" & _\n' +
			'        "world"\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'invalid-line-continuation')).toHaveLength(0);
	});

	it('reads an underscore with whitespace after it as a continuation, as the VBE does (issue #83)', () => {
		const src =
			'Option Explicit\n' +
			'Sub T()\n' +
			'    Dim a As Boolean, _ \n' +
			'        b As Boolean\n' +
			'    If a And _  \n' +
			'       b Then\n' +
			'        Debug.Print "both"\n' +
			'    End If\n' +
			'End Sub\n';

		expect(analyzeModule(src).map((d) => d.code)).toEqual([]);
	});

	it('ignores underscores in identifiers, strings, and comments', () => {
		const src =
			'Sub T()\n' +
			'    Dim value_name As Long\n' +
			'    value_name = 1\n' +
			'    Debug.Print "text _ inside string"\n' +
			"    ' comment _ at the end\n" +
			'    Rem comment _ at the end\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'invalid-line-continuation')).toHaveLength(0);
	});

	it('leaves a final dangling continuation as a realtime recovery case', () => {
		const src = 'Sub T()\n    Debug.Print "hello" & _';

		expect(byCode(analyzeModule(src), 'invalid-line-continuation')).toHaveLength(0);
	});

	it('stays quiet when surrounding code is unparseable but every string is closed', () => {
		// No-FP control: token-recovery / malformed surroundings must not trip
		// unterminated-string when no string literal is actually left open.
		const src =
			'Sub T()\n' +
			'    Dim x As\n' +
			'    x = ((1 +\n' +
			'    MsgBox "say ""hi"" now"\n' +
			'    Debug.Print "ok" & @#$\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'unterminated-string')).toHaveLength(0);
	});
});
