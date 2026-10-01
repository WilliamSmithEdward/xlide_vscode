// Diagnostics tests: a line label cannot be a reserved word (issue #272).
// Every reserved word was measured in Excel 16.0 (build 20326, 2026-10-01)
// by compiling the project, as `GoTo <word>` then `<word>:`, and as the
// label alone.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

function source(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.join('\n')}\n    Main = 1\nEnd Function\n`;
}

function malformed(src: string): string[] {
	return byCode(analyzeModule(src), 'malformed-statement').map((diag) => src.slice(diag.span.start, diag.span.end));
}

describe('a reserved word as a line label', () => {
	it.each(['Fix', 'Int', 'Len', 'LenB', 'Abs', 'Sgn', 'CStr', 'CDate', 'Array', 'String', 'Debug', 'Erase', 'Dim', 'Const', 'Long', 'True', 'Local'])('reports GoTo %s', (word) => {
		const src = source(`    GoTo ${word}`, `${word}:`);
		const found = malformed(src);
		expect(found).toContain(word);
		const diag = byCode(analyzeModule(src), 'malformed-statement').find((d) => d.message.includes('jump to'));
		expect(diag?.message).toBe(`'${word}' is a reserved word, so it cannot name a line label to jump to. This is a VBE compile error: Syntax error.`);
	});

	it.each(['Fix', 'Debug', 'Dim', 'Const', 'Static', 'Enum', 'Type', 'Call', 'Then', 'Nothing'])('reports the label %s: alone', (word) => {
		const src = source(`${word}:`);
		expectDiagnostic(src, byCode(analyzeModule(src), 'malformed-statement'), 'malformed-statement', { span: word, message: `'${word}' is a reserved word, so '${word}:' cannot be a line label. This is a VBE compile error: Syntax error.` });
	});

	it.each([
		['Print', 'Method not valid without suitable object'],
		['Scale', 'Method not valid without suitable object'],
		['CDec', 'Argument not optional'],
		['Date', 'Invalid use of property'],
	])('reports %s: as the statement it is', (word, error) => {
		const src = source(`${word}:`);
		expectDiagnostic(src, byCode(analyzeModule(src), 'malformed-statement'), 'malformed-statement', { span: word, message: `'${word}:' is the statement ${word}, not a line label. This is a VBE compile error: ${error}.` });
	});

	it('reports each jump form', () => {
		for (const line of ['    GoSub Fix', '    On Error GoTo Fix', '    Resume Fix', '    On 1 GoTo Here, Fix']) {
			expect(malformed(source(line, 'Here:')), line).toEqual(['Fix']);
		}
	});

	it('stays quiet on a library name a label may take, and the statements a reserved word starts', () => {
		for (const word of ['Mid', 'Left', 'Error', 'UCase', 'Now', 'Sqr', 'Split', 'InStr', 'MsgBox', 'Object', 'Name', 'Kill']) {
			expect(malformed(source(`    GoTo ${word}`, `${word}:`)), word).toEqual([]);
		}
		for (const line of ['    Stop: Main = 2', '    DoEvents: Main = 2', '    Close: Main = 2', '    On Error GoTo 0', '    On Error Resume Next: Resume Next']) {
			expect(malformed(source(line)), line).toEqual([]);
		}
		expect(malformed(source('    If Main = 1 Then', '        Main = 2', '    Else: Main = 3', '    End If'))).toEqual([]);
		// A continued line is no label: `Debug.Print _` then `Date: ...`.
		expect(malformed(source('    Debug.Print _', '        Date: Main = 2'))).toEqual([]);
	});
});
