// Characters and lines the VBE refuses (issues #132 and #133). Each refused
// sample was measured in Excel 16.0 (build 20326, 2026-09-26); each accepted
// one compiles there.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic, expectDiagnostics } from '../helpers/diagnostics';

const CODE = 'stray-character';
const NBSP = '\u00A0';

function wrap(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n    Dim n As Long\n${lines.map((line) => `    ${line}`).join('\n')}\n    Main = n\nEnd Function\n`;
}

describe('stray-character (issue #132)', () => {
	it('flags a trailing semicolon and the characters VBA does not use', () => {
		expectDiagnostic(wrap('n = 1;'), analyzeModule(wrap('n = 1;')), CODE, { span: ';', message: 'Print or Write' });
		for (const [line, span] of [['n = 1 `', '`'], ['n = @', '@'], ['n = ~1', '~'], ['n = 1 | 2', '|']]) {
			const src = wrap(line);
			expectDiagnostics(src, analyzeModule(src), CODE, [{ span, message: 'Syntax error' }]);
		}
		const braces = wrap('n = {1}');
		expectDiagnostics(braces, analyzeModule(braces), CODE, [{ span: '{' }, { span: '}' }]);
	});

	it('flags a non-breaking space between tokens and one used as indentation, once per run', () => {
		const between = wrap(`n = 1${NBSP}+${NBSP}1`);
		expectDiagnostics(between, analyzeModule(between), CODE, [{ message: 'U+00A0' }, { message: 'U+00A0' }]);
		const indent = `Option Explicit\nFunction Main() As Variant\n    Dim n As Long\n${NBSP}${NBSP}${NBSP}${NBSP}n = 1\n    Main = n\nEnd Function\n`;
		expectDiagnostics(indent, analyzeModule(indent), CODE, [{ span: NBSP.repeat(1), message: 'non-breaking space' }]);
	});

	it('leaves Print lists, type-declaration suffixes and directives alone', () => {
		const src = wrap('Debug.Print n; n', 'Debug.Print n,', 'Dim s$, i%, l&, d#, f!, c@', 's = Left$("ab", 1)', 'Print #1, n; n');
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});
});

describe('line-too-long (issue #133)', () => {
	it('refuses a physical line of 1024 characters and accepts 1023', () => {
		const comment = (length: number): string => `    n = 1 '${'x'.repeat(length - '    n = 1 \''.length)}`;
		const ok = wrap(comment(1019)); // four spaces of indent make 1023
		expect(byCode(analyzeModule(ok), 'line-too-long')).toHaveLength(0);
		const long = wrap(comment(1020));
		expectDiagnostic(long, analyzeModule(long), 'line-too-long', { message: ['1024 characters', '1023'] });
	});
});
