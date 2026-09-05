// Inline Variable (issue #69). The interesting refusal is VBA-specific: every
// other language inlines a compound value by bracketing it to keep precedence,
// but in VBA `Foo (x)` passes x by value where `Foo x` passes it by reference,
// so there is no bracket that is safe to add. A compound value is refused.

import { describe, expect, it } from 'vitest';
import { inlineVariable } from '../src/analyzer/refactor/inlineVariable';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';

function run(source: string, name: string) {
	const offset = source.indexOf(name);
	expect(offset, `'${name}' is not in the fixture`).toBeGreaterThan(-1);
	return inlineVariable({ source, offset });
}

function applied(source: string, name: string): string {
	const result = run(source, name);
	if (!result.ok) { throw new Error(`refused: ${result.reason}`); }
	return applyVbaTextEdits(source, result.edits);
}

function reason(source: string, name: string): string {
	const result = run(source, name);
	if (result.ok) { throw new Error(`expected a refusal, got ${result.title}`); }
	return result.reason;
}

describe('what it writes', () => {
	it('drops the declaration and the assignment and substitutes the value', () => {
		const source = [
			'Public Sub Go()',
			'    Dim limit As Long',
			'    limit = 3',
			'    Debug.Print limit',
			'End Sub',
			'',
		].join('\r\n');
		expect(applied(source, 'limit')).toBe([
			'Public Sub Go()',
			'    Debug.Print 3',
			'End Sub',
			'',
		].join('\r\n'));
	});

	it('substitutes at every use', () => {
		const source = [
			'Public Sub Go()',
			'    Dim limit As Long',
			'    limit = 3',
			'    Debug.Print limit',
			'    Debug.Print limit + limit',
			'End Sub',
			'',
		].join('\r\n');
		const out = applied(source, 'limit');
		expect(out).toContain('    Debug.Print 3\r\n');
		expect(out).toContain('    Debug.Print 3 + 3\r\n');
		expect(out).not.toContain('limit');
	});

	it('inlines a string, a date, a hex literal and a keyword value', () => {
		for (const [value, used] of [
			['"text"', '"text"'],
			['#2026-01-01#', '#2026-01-01#'],
			['&HFF&', '&HFF&'],
			['True', 'True'],
			['Nothing', 'Nothing'],
		] as const) {
			const source = [
				'Public Sub Go()',
				'    Dim v As Variant',
				`    v = ${value}`,
				'    Debug.Print v',
				'End Sub',
				'',
			].join('\r\n');
			expect(applied(source, 'v'), value).toContain(`Debug.Print ${used}`);
		}
	});

	it('inlines a plain name, which needs no brackets either', () => {
		const source = [
			'Public Sub Go(ByVal n As Long)',
			'    Dim copy As Long',
			'    copy = n',
			'    Debug.Print copy',
			'End Sub',
			'',
		].join('\r\n');
		expect(applied(source, 'copy')).toContain('    Debug.Print n\r\n');
	});

	it('works from a use as well as from the declaration', () => {
		const source = [
			'Public Sub Go()',
			'    Dim limit As Long',
			'    limit = 3',
			'    Debug.Print limit',
			'End Sub',
			'',
		].join('\r\n');
		const atUse = source.lastIndexOf('limit');
		const result = inlineVariable({ source, offset: atUse });
		expect(result.ok).toBe(true);
		expect(result.ok && applyVbaTextEdits(source, result.edits)).toContain('Debug.Print 3');
	});
});

describe('what it refuses, and why', () => {
	function local(body: string[]): string {
		return ['Public Sub Go()', ...body, 'End Sub', ''].join('\r\n');
	}

	it('refuses a compound value rather than bracketing it', () => {
		const source = local(['    Dim total As Long', '    total = 1 + 2', '    Debug.Print total']);
		expect(reason(source, 'total')).toMatch(/ByRef arguments to ByVal/);
	});

	it('refuses a call, which is a compound value too', () => {
		const source = local(['    Dim n As Long', '    n = Len("ab")', '    Debug.Print n']);
		expect(reason(source, 'n')).toMatch(/not a single value/);
	});

	it('refuses a variable assigned more than once', () => {
		const source = local([
			'    Dim limit As Long',
			'    limit = 3',
			'    limit = 4',
			'    Debug.Print limit',
		]);
		expect(reason(source, 'limit')).toMatch(/assigned 2 times/);
	});

	it('refuses a variable that is never assigned', () => {
		const source = local(['    Dim limit As Long', '    Debug.Print limit']);
		expect(reason(source, 'limit')).toMatch(/never assigned/);
	});

	it('refuses a read that happens before the assignment', () => {
		const source = local([
			'    Dim limit As Long',
			'    Debug.Print limit',
			'    limit = 3',
			'    Debug.Print limit',
		]);
		expect(reason(source, 'limit')).toMatch(/read before it is assigned/);
	});

	it('refuses Static, which keeps its value between calls', () => {
		const source = local(['    Static limit As Long', '    limit = 3', '    Debug.Print limit']);
		expect(reason(source, 'limit')).toMatch(/Static/);
	});

	it('refuses an array, which has no single value', () => {
		const source = local(['    Dim items(1 To 3) As Long', '    Debug.Print items(1)']);
		expect(reason(source, 'items')).toMatch(/array/);
	});

	it('refuses a shared declaration line', () => {
		const source = local(['    Dim a As Long, b As Long', '    a = 1', '    Debug.Print a']);
		expect(reason(source, 'a')).toMatch(/shares its declaration line/);
	});

	it('refuses a module variable, which is not a local', () => {
		const source = 'Public Total As Long\r\n\r\nPublic Sub Go()\r\n    Total = 1\r\nEnd Sub\r\n';
		expect(reason(source, 'Total')).toMatch(/inside a procedure|not a local/);
	});

	it('refuses a Const, which already is its value', () => {
		const source = local(['    Const limit As Long = 3', '    Debug.Print limit']);
		expect(reason(source, 'limit')).toMatch(/Const/);
	});
});
