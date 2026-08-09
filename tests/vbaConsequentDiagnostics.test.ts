import { describe, expect, it } from 'vitest';
import { analyzeVbaModuleSource } from '../src/vbaModuleAnalysis';
import { matchOpener } from '../src/vbaStructuralDiagnostics';

// One malformed procedure header used to produce two findings: the real one,
// and a consequence of it. Block matching refused to see the header, so the
// block never opened and its closer looked orphaned - a second error naming a
// line where nothing is wrong. The analyzer is shared with xlide_vbide, so the
// noise showed up identically in both products.
function codesFor(lines: string[]): string[] {
	return analyzeVbaModuleSource({
		source: lines.join('\r\n'),
		moduleName: 'Module1',
		moduleType: 'standard',
		moduleKind: 'standard',
	}).diagnostics.map((d) => String(d.code));
}

describe('a malformed header does not also report its closer', () => {
	it('reports only the invalid name for a bad Sub header', () => {
		const codes = codesFor(['Option Explicit', '', 'Public Sub 1Bad()', '    Debug.Print "x"', 'End Sub', '']);
		expect(codes).toContain('invalid-identifier-start');
		expect(codes).not.toContain('unmatched-block-closer');
	});

	it('reports only the invalid name for a bad Function header', () => {
		const codes = codesFor(['Option Explicit', '', 'Public Function 9Bad() As Long', '    9Bad = 1', 'End Function', '']);
		expect(codes).toContain('invalid-identifier-start');
		expect(codes).not.toContain('unmatched-block-closer');
	});

	it.each([
		['Type', ['Option Explicit', '', 'Public Type 1Bad', '    X As Long', 'End Type', '']],
		['Enum', ['Option Explicit', '', 'Public Enum 2Bad', '    First = 1', 'End Enum', '']],
		['Property', ['Option Explicit', '', 'Public Property Get 3Bad() As Long', '    3Bad = 1', 'End Property', '']],
	])('does not report an orphan closer for a bad %s header', (_kind, lines) => {
		expect(codesFor(lines as string[])).not.toContain('unmatched-block-closer');
	});

	// The counter-case the issue calls out: a closer with no opener anywhere is
	// a real finding and has to survive.
	it('still reports a genuinely orphaned closer', () => {
		expect(codesFor(['Option Explicit', '', 'End Sub', ''])).toContain('unmatched-block-closer');
	});

	it('still reports a missing closer', () => {
		expect(codesFor(['Option Explicit', '', 'Public Sub Fine()', '    Debug.Print "x"', '']))
			.toContain('missing-block-closer');
	});

	it('leaves valid code clean', () => {
		expect(codesFor(['Option Explicit', '', 'Public Sub Good()', '    Debug.Print "x"', 'End Sub', ''])).toEqual([]);
	});
});

describe('matchOpener sees a header whose name is invalid', () => {
	it.each([
		['Public Sub 1Bad()', 'Sub', 'Sub 1Bad'],
		['Public Function 9Bad() As Long', 'Function', 'Function 9Bad'],
		['Public Type 1Bad', 'Type', 'Type 1Bad'],
		['Public Enum 2Bad', 'Enum', 'Enum 2Bad'],
	])('%s opens a block', (line, kind, label) => {
		const opener = matchOpener(line);
		expect(opener?.kind).toBe(kind);
		expect(opener?.label).toBe(label);
	});

	it.each([
		'End Sub',
		'Sub',
		'Public Sub',
		'Dim Subtotal As Long',
		'Set Functionality = Nothing',
	])('%s is not a block opener', (line) => {
		expect(matchOpener(line)).toBeUndefined();
	});
});
