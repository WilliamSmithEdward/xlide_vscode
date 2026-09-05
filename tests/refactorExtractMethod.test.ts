// Extract Method (issue #69). The signature is read off the analyzer's
// reference kinds (#55) rather than guessed: what the selection reads before
// writing is a parameter, what it writes and the caller reads afterwards is
// the result, and what it writes and nobody reads afterwards takes its Dim
// with it.

import { describe, expect, it } from 'vitest';
import { extractMethod } from '../src/analyzer/refactor/extractMethod';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';

/**
 * The fixture marks the selection with `<<` and `>>` on their own lines. They
 * are removed, and the span is the lines that were between them.
 */
function run(fixture: string, name = 'Extracted') {
	const eol = fixture.includes('\r\n') ? '\r\n' : '\n';
	const lines = fixture.split(eol);
	const openAt = lines.indexOf('<<');
	const closeAt = lines.indexOf('>>');
	expect(openAt, 'fixture needs a << line').toBeGreaterThan(-1);
	expect(closeAt, 'fixture needs a >> line').toBeGreaterThan(openAt);

	const kept = lines.filter((line) => line !== '<<' && line !== '>>');
	const source = kept.join(eol);
	const offsetOf = (index: number): number =>
		kept.slice(0, index).reduce((total, line) => total + line.length + eol.length, 0);

	// Removing `<<` shifts the selected lines down one: they are kept[openAt]
	// through kept[closeAt - 2].
	const span = { start: offsetOf(openAt), end: offsetOf(closeAt - 1) - eol.length };
	return { source, result: extractMethod({ source, span, name }) };
}

function applied(fixture: string, name?: string): string {
	const { source, result } = run(fixture, name);
	if (!result.ok) { throw new Error(`refused: ${result.reason}`); }
	return applyVbaTextEdits(source, result.edits);
}

function reason(fixture: string, name?: string): string {
	const { result } = run(fixture, name);
	if (result.ok) { throw new Error(`expected a refusal, got ${result.title}`); }
	return result.reason;
}

describe('what it writes', () => {
	it('extracts statements that need nothing into a bare Sub', () => {
		const out = applied([
			'Option Explicit',
			'',
			'Public Sub Go()',
			'<<',
			'    Debug.Print "one"',
			'    Debug.Print "two"',
			'>>',
			'End Sub',
			'',
		].join('\r\n'));
		expect(out).toContain('    Extracted\r\n');
		expect(out).toContain('Private Sub Extracted()');
		expect(out).toContain('    Debug.Print "one"');
		expect(out).toContain('End Sub');
	});

	it('passes ByVal what the selection reads and never writes', () => {
		const out = applied([
			'Option Explicit',
			'',
			'Public Sub Go()',
			'    Dim n As Long',
			'    n = 3',
			'<<',
			'    Debug.Print n',
			'>>',
			'End Sub',
			'',
		].join('\r\n'));
		expect(out).toContain('Private Sub Extracted(ByVal n As Long)');
		expect(out).toContain('    Extracted n\r\n');
	});

	it('returns a value the caller reads afterwards', () => {
		const out = applied([
			'Option Explicit',
			'',
			'Public Sub Go()',
			'    Dim total As Long',
			'<<',
			'    total = 3',
			'>>',
			'    Debug.Print total',
			'End Sub',
			'',
		].join('\r\n'));
		expect(out).toContain('Private Function Extracted() As Long');
		expect(out).toContain('    Extracted = total');
		expect(out).toContain('    total = Extracted()');
		expect(out).toContain('End Function');
	});

	it('passes ByRef what it both reads and writes and the caller reads after', () => {
		const out = applied([
			'Option Explicit',
			'',
			'Public Sub Go()',
			'    Dim total As Long',
			'    total = 1',
			'<<',
			'    total = total + 1',
			'>>',
			'    Debug.Print total',
			'End Sub',
			'',
		].join('\r\n'));
		expect(out).toContain('Private Sub Extracted(ByRef total As Long)');
		expect(out).toContain('    Extracted total\r\n');
	});

	it('takes the Dim with it when nothing after reads the variable', () => {
		const out = applied([
			'Option Explicit',
			'',
			'Public Sub Go()',
			'    Dim scratch As Long',
			'<<',
			'    scratch = 3',
			'    Debug.Print scratch',
			'>>',
			'End Sub',
			'',
		].join('\r\n'));
		expect(out).toContain('Private Sub Extracted()');
		expect(out).toContain('    Dim scratch As Long');
		// And the caller no longer declares it.
		expect(out.slice(0, out.indexOf('Private Sub Extracted')))
			.not.toContain('Dim scratch');
	});

	it('needs nothing for a name that is not a local', () => {
		const out = applied([
			'Option Explicit',
			'',
			'Public Total As Long',
			'',
			'Public Sub Go()',
			'<<',
			'    Total = Total + 1',
			'>>',
			'End Sub',
			'',
		].join('\r\n'));
		expect(out).toContain('Private Sub Extracted()');
		expect(out).toContain('    Extracted\r\n');
	});

	it('puts the new procedure below the one it came out of', () => {
		const out = applied([
			'Option Explicit',
			'',
			'Public Sub Go()',
			'<<',
			'    Debug.Print 1',
			'>>',
			'End Sub',
			'',
		].join('\r\n'));
		expect(out.indexOf('Private Sub Extracted')).toBeGreaterThan(out.indexOf('Public Sub Go'));
	});

	it('uses the name it is given', () => {
		const out = applied([
			'Option Explicit',
			'',
			'Public Sub Go()',
			'<<',
			'    Debug.Print 1',
			'>>',
			'End Sub',
			'',
		].join('\r\n'), 'PrintOne');
		expect(out).toContain('Private Sub PrintOne()');
		expect(out).toContain('    PrintOne\r\n');
	});
});

describe('what it refuses, and why', () => {
	it('refuses a module without Option Explicit', () => {
		expect(reason([
			'Public Sub Go()',
			'<<',
			'    Debug.Print 1',
			'>>',
			'End Sub',
			'',
		].join('\r\n'))).toMatch(/Option Explicit/);
	});

	it('refuses a Static local, which keeps its value between calls', () => {
		expect(reason([
			'Option Explicit',
			'',
			'Public Sub Go()',
			'    Static seen As Long',
			'<<',
			'    seen = seen + 1',
			'>>',
			'End Sub',
			'',
		].join('\r\n'))).toMatch(/Static/);
	});

	it('refuses a name the module already has', () => {
		expect(reason([
			'Option Explicit',
			'',
			'Public Sub Go()',
			'<<',
			'    Debug.Print 1',
			'>>',
			'End Sub',
			'',
			'Private Sub Extracted()',
			'End Sub',
			'',
		].join('\r\n'))).toMatch(/already has a procedure called 'Extracted'/);
	});

	it('refuses a selection that reaches outside one procedure', () => {
		expect(reason([
			'Option Explicit',
			'',
			'Public Sub One()',
			'<<',
			'    Debug.Print 1',
			'End Sub',
			'',
			'Public Sub Two()',
			'    Debug.Print 2',
			'>>',
			'End Sub',
			'',
		].join('\r\n'))).toMatch(/inside one procedure/);
	});

	it('refuses a selection that takes the End line with it', () => {
		expect(reason([
			'Option Explicit',
			'',
			'Public Sub Go()',
			'<<',
			'    Debug.Print 1',
			'End Sub',
			'>>',
			'',
		].join('\r\n'))).toMatch(/header or End line/);
	});
});
