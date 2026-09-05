// Extract Variable (issue #69): a selected expression is declared and assigned
// above its statement and replaced by the name. The type and the `Set` come
// from resolveExpressionType (#61), so this and the declare-variable quick fix
// never disagree about what an expression is.

import { describe, expect, it } from 'vitest';
import { extractVariable } from '../src/analyzer/refactor/extractVariable';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';

/** Extract the first occurrence of `expression` in the fixture. */
function run(source: string, expression: string) {
	const start = source.indexOf(expression);
	expect(start, `'${expression}' is not in the fixture`).toBeGreaterThan(-1);
	return extractVariable({ source, span: { start, end: start + expression.length } });
}

function applied(source: string, expression: string): string {
	const result = run(source, expression);
	if (!result.ok) { throw new Error(`refused: ${result.reason}`); }
	return applyVbaTextEdits(source, result.edits);
}

function reason(source: string, expression: string): string {
	const result = run(source, expression);
	if (result.ok) { throw new Error(`expected a refusal, got ${result.title}`); }
	return result.reason;
}

describe('what it writes', () => {
	it('declares above the statement and replaces the selection', () => {
		const source = [
			'Option Explicit',
			'',
			'Public Sub Go()',
			'    Debug.Print 2 * 3',
			'End Sub',
			'',
		].join('\r\n');
		expect(applied(source, '2 * 3')).toBe([
			'Option Explicit',
			'',
			'Public Sub Go()',
			'    Dim value As Double',
			'    value = 2 * 3',
			'    Debug.Print value',
			'End Sub',
			'',
		].join('\r\n'));
	});

	it('keeps the statement\'s own indentation', () => {
		const source = [
			'Public Sub Go()',
			'    If True Then',
			'            Debug.Print 2 * 3',
			'    End If',
			'End Sub',
			'',
		].join('\r\n');
		const out = applied(source, '2 * 3');
		expect(out).toContain('\r\n            Dim value As Double\r\n');
		expect(out).toContain('\r\n            value = 2 * 3\r\n');
	});

	it('names the variable after the last name in the expression', () => {
		const source = 'Public Sub Go()\r\n    Debug.Print ActiveSheet.Name\r\nEnd Sub\r\n';
		const out = applied(source, 'ActiveSheet.Name');
		// Variant is the analyzer's answer for ActiveSheet.Name without project
		// context, and it is a real answer: it is what VBA gives the name too.
		expect(out).toContain('Dim name As Variant');
		expect(out).toContain('    name = ActiveSheet.Name');
		expect(out).toContain('    Debug.Print name');
	});

	it('assigns an object with Set', () => {
		const source = 'Public Sub Go()\r\n    Debug.Print ActiveSheet.Range("A1").Address\r\nEnd Sub\r\n';
		const out = applied(source, 'ActiveSheet.Range("A1")');
		expect(out).toContain('Dim range As Excel.Range');
		expect(out).toContain('    Set range = ActiveSheet.Range("A1")');
		// Not 'a1': the name comes from the expression, and "A1" is a string.
		expect(out).not.toContain('Dim a1');
	});

	it('does not shadow a name the procedure already uses', () => {
		const source = [
			'Public Sub Go()',
			'    Dim value As Long',
			'    Debug.Print 2 * 3',
			'End Sub',
			'',
		].join('\r\n');
		const out = applied(source, '2 * 3');
		expect(out).toContain('Dim value2 As Double');
		expect(out).toContain('    value2 = 2 * 3');
	});

	it('does not shadow a parameter or a module variable either', () => {
		const source = [
			'Public value As Long',
			'',
			'Public Sub Go(ByVal value2 As Long)',
			'    Debug.Print 2 * 3',
			'End Sub',
			'',
		].join('\r\n');
		expect(applied(source, '2 * 3')).toContain('Dim value3 As Double');
	});

	it('falls back to a plain name when the expression names nothing usable', () => {
		const source = 'Public Sub Go()\r\n    Debug.Print True And False\r\nEnd Sub\r\n';
		expect(applied(source, 'True And False')).toContain('Dim value As ');
	});

	it('points the caret at the declared name, which is what gets renamed', () => {
		const source = 'Public Sub Go()\r\n    Debug.Print 2 * 3\r\nEnd Sub\r\n';
		const result = run(source, '2 * 3');
		if (!result.ok) { throw new Error(result.reason); }
		const applied = applyVbaTextEdits(source, result.edits);
		expect(applied.slice(result.renameSpan!.start, result.renameSpan!.end)).toBe('value');
	});

	it('keeps the module\'s own line endings', () => {
		const out = applied('Public Sub Go()\n    Debug.Print 2 * 3\nEnd Sub\n', '2 * 3');
		expect(out).toContain('\n    value = 2 * 3\n');
		expect(out).not.toContain('\r');
	});
});

describe('what it refuses, and why', () => {
	it('refuses an empty selection', () => {
		const source = 'Public Sub Go()\r\n    Debug.Print 1\r\nEnd Sub\r\n';
		const result = extractVariable({ source, span: { start: 20, end: 20 } });
		expect(result.ok).toBe(false);
		expect(!result.ok && result.reason).toMatch(/Select an expression/);
	});

	it('refuses a selection outside any procedure', () => {
		const source = 'Public Const Limit As Long = 3\r\n\r\nPublic Sub Go()\r\nEnd Sub\r\n';
		expect(reason(source, '3')).toMatch(/inside a procedure/);
	});

	it('refuses half an expression, which would not parse on its own', () => {
		const source = 'Public Sub Go()\r\n    Debug.Print 2 * 3\r\nEnd Sub\r\n';
		expect(reason(source, '2 *')).toMatch(/part of an expression/);
	});

	it('refuses a selection that crosses two statements', () => {
		const source = [
			'Public Sub Go()',
			'    Debug.Print 1',
			'    Debug.Print 2',
			'End Sub',
			'',
		].join('\r\n');
		expect(reason(source, '1\r\n    Debug.Print 2')).toMatch(/inside a single statement/);
	});

	it('refuses inside a declaration, which VBA cannot initialise', () => {
		const source = 'Public Sub Go()\r\n    Dim n As Long\r\nEnd Sub\r\n';
		expect(reason(source, 'Long')).toMatch(/declaration has nothing to extract/);
	});
});
