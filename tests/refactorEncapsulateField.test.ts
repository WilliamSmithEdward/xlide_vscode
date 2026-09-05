// Encapsulate Field (issue #69): a public module variable becomes private
// behind a property pair that keeps its name, so no call site is rewritten.
// The refusals are the point - each one names a shape where the rewrite would
// change what the code means, or would not compile.

import { describe, expect, it } from 'vitest';
import { encapsulateField } from '../src/analyzer/refactor/encapsulateField';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';

/** Encapsulate the field the given name declares, and return the new module. */
function run(source: string, name: string, projectClassNames?: readonly string[]) {
	const offset = source.indexOf(name);
	expect(offset, `'${name}' is not in the fixture`).toBeGreaterThan(-1);
	return encapsulateField({ source, offset, ...(projectClassNames ? { projectClassNames } : {}) });
}

function applied(source: string, name: string, projectClassNames?: readonly string[]): string {
	const result = run(source, name, projectClassNames);
	if (!result.ok) { throw new Error(`refused: ${result.reason}`); }
	return applyVbaTextEdits(source, result.edits);
}

function reason(source: string, name: string): string {
	const result = run(source, name);
	if (result.ok) { throw new Error('expected a refusal'); }
	return result.reason;
}

describe('what it writes', () => {
	it('keeps the name on the property and gives the field an m_ prefix', () => {
		const source = 'Option Explicit\r\nPublic Total As Long\r\n';
		expect(applied(source, 'Total')).toBe([
			'Option Explicit',
			'Private m_Total As Long',
			'',
			'Public Property Get Total() As Long',
			'    Total = m_Total',
			'End Property',
			'',
			'Public Property Let Total(ByVal RHS As Long)',
			'    m_Total = RHS',
			'End Property',
			'',
		].join('\r\n'));
	});

	it('uses Property Set, and Set on both assignments, for an object type', () => {
		const source = 'Option Explicit\r\nPublic Sheet As Worksheet\r\n';
		const out = applied(source, 'Sheet');
		expect(out).toContain('Public Property Set Sheet(ByVal RHS As Worksheet)');
		expect(out).toContain('    Set Sheet = m_Sheet');
		expect(out).toContain('    Set m_Sheet = RHS');
		expect(out).not.toContain('Property Let');
	});

	it('treats a project class as an object', () => {
		const out = applied('Public Widget As Thing\r\n', 'Widget', ['Thing']);
		expect(out).toContain('Public Property Set Widget(ByVal RHS As Thing)');
	});

	it('writes a type suffix out in full, so the field matches its property', () => {
		const out = applied('Public Count%\r\n', 'Count');
		expect(out).toContain('Private m_Count As Integer');
		expect(out).toContain('Public Property Get Count() As Integer');
		expect(out).toContain('Public Property Let Count(ByVal RHS As Integer)');
	});

	it('gives an untyped variable the Variant it already has', () => {
		const out = applied('Public Thing\r\n', 'Thing');
		expect(out).toContain('Public Property Get Thing() As Variant');
		expect(out).toContain('Public Property Let Thing(ByVal RHS As Variant)');
	});

	it('keeps a fixed-length String fixed on the backing field', () => {
		const out = applied('Public Code As String * 8\r\n', 'Code');
		expect(out).toContain('Private m_Code As String * 8');
		// The property itself cannot be fixed-length; it returns a String.
		expect(out).toContain('Public Property Get Code() As String');
	});

	it('keeps the module\'s own line endings', () => {
		expect(applied('Public Total As Long\n', 'Total')).toContain('End Property\n');
	});

	it('titles the action with the field name', () => {
		const result = run('Public Total As Long\r\n', 'Total');
		expect(result.ok && result.title).toBe("Encapsulate 'Total' behind a property");
	});
});

describe('what it refuses, and why', () => {
	it('refuses a Const, which has no value to set', () => {
		expect(reason('Public Const Limit As Long = 3\r\n', 'Limit')).toMatch(/Const/);
	});

	it('refuses a Private variable, which nothing outside reads', () => {
		expect(reason('Private Total As Long\r\n', 'Total')).toMatch(/already Private/);
	});

	it('refuses WithEvents, whose events a property cannot raise', () => {
		expect(reason('Public WithEvents App As Application\r\n', 'App')).toMatch(/WithEvents/);
	});

	it('refuses an array, which VBA cannot return by reference', () => {
		expect(reason('Public Items(1 To 10) As Long\r\n', 'Items')).toMatch(/array/);
	});

	it('refuses a shared declaration line, naming what it shares with', () => {
		expect(reason('Public Total As Long, Count As Long\r\n', 'Total'))
			.toMatch(/shares its declaration with 'Count'/);
	});

	it('refuses a declaration continued across lines', () => {
		expect(reason('Public Total _\r\n    As Long\r\n', 'Total')).toMatch(/continued across lines/);
	});

	it('refuses when the backing name is taken', () => {
		expect(reason('Public Total As Long\r\nPrivate m_Total As Long\r\n', 'Total'))
			.toMatch(/already has something called 'm_Total'/);
	});

	it('refuses when a procedure already has the property name', () => {
		const source = 'Public Total As Long\r\n\r\nPublic Sub Total()\r\nEnd Sub\r\n';
		expect(reason(source, 'Total')).toMatch(/already has a procedure called 'Total'/);
	});

	it('refuses a caret that is not on a module variable at all', () => {
		const source = 'Public Sub Go()\r\n    Dim local As Long\r\nEnd Sub\r\n';
		expect(reason(source, 'local')).toMatch(/module-level variable/);
	});
});
