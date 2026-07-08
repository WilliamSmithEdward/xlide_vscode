import { describe, expect, it } from 'vitest';
import { analyzeVbaModuleSource } from '../src/vbaModuleAnalysis';

// Regression suite from analyzing the real-world ModernJsonInVBA library:
// five families of valid VBA that were wrongly flagged as errors.

function codes(body: string): string[] {
	const src = `Option Explicit\n${body}`;
	// knownIdentifiers activates the undeclared-variable rule (it is inert
	// without a known-identifier universe, mirroring the project pipeline).
	return analyzeVbaModuleSource({ source: src, moduleName: 'Module1', knownIdentifiers: new Set<string>() })
		.diagnostics.map((d) => `${d.code}:${/'([^']+)'/.exec(d.message)?.[1] ?? ''}`);
}

describe('Open-statement Access clause', () => {
	it.each([
		'Access Read',
		'Access Write',
		'Access Read Write',
	])('does not flag Access in: Open ... For Binary %s As #f', (clause) => {
		const body = `Sub T()\n    Dim f As Long\n    Dim p As String\n    f = FreeFile\n    Open p For Binary ${clause} As #f\n    Close #f\nEnd Sub\n`;
		expect(codes(body)).not.toContain('undeclared-variable:Access');
	});

	it('still flags an undeclared variable named Access outside the clause position', () => {
		const body = 'Sub T()\n    Dim x As Long\n    x = Access\nEnd Sub\n';
		expect(codes(body)).toContain('undeclared-variable:Access');
	});

	it('still flags an undeclared pathname inside an Open statement', () => {
		const body = 'Sub T()\n    Dim f As Long\n    f = FreeFile\n    Open missingPath For Binary Access Read As #f\nEnd Sub\n';
		expect(codes(body)).toContain('undeclared-variable:missingPath');
	});
});

describe('Byte array assigned to a String scalar', () => {
	it('does not flag s = byteArray (documented VBA conversion)', () => {
		const body = 'Sub T()\n    Dim o() As Byte\n    Dim s As String\n    ReDim o(0 To 3)\n    s = o\nEnd Sub\n';
		expect(codes(body).some((c) => c.startsWith('array-assignment-to-scalar'))).toBe(false);
	});

	it('does not flag a String function return assigned from a Byte array', () => {
		const body = 'Function F() As String\n    Dim o() As Byte\n    ReDim o(0 To 3)\n    F = o\nEnd Function\n';
		expect(codes(body).some((c) => c.startsWith('array-assignment-to-scalar'))).toBe(false);
	});

	it.each([
		['Long array to String', 'Sub T()\n    Dim o() As Long\n    Dim s As String\n    ReDim o(0 To 3)\n    s = o\nEnd Sub\n'],
		['Byte array to Long', 'Sub T()\n    Dim o() As Byte\n    Dim n As Long\n    ReDim o(0 To 3)\n    n = o\nEnd Sub\n'],
	])('still flags %s', (_label, body) => {
		expect(codes(body).some((c) => c.startsWith('array-assignment-to-scalar'))).toBe(true);
	});
});

describe('ReDim inside a single-line If', () => {
	it('does not flag the ReDim target as an unallocated access', () => {
		const body = 'Sub T()\n    Dim headers() As String\n    Dim colCount As Long\n    colCount = 3\n    If colCount > 0 Then ReDim headers(1 To colCount)\nEnd Sub\n';
		expect(codes(body).some((c) => c.startsWith('unallocated-dynamic-array-access'))).toBe(false);
	});

	it('does not flag either arm of If ... Then ReDim ... Else ReDim ...', () => {
		const body = 'Sub T()\n    Dim a() As String\n    Dim c As Long\n    c = 2\n    If c > 0 Then ReDim a(1 To c) Else ReDim a(1 To 5)\nEnd Sub\n';
		expect(codes(body).some((c) => c.startsWith('unallocated-dynamic-array-access'))).toBe(false);
	});

	it('still flags a genuine unallocated access, including inside a single-line If', () => {
		const noRedim = 'Sub T()\n    Dim a() As String\n    Dim x As String\n    x = a(1)\nEnd Sub\n';
		const ifAccess = 'Sub T()\n    Dim a() As String\n    Dim x As String\n    Dim c As Long\n    c = 1\n    If c > 0 Then x = a(1)\nEnd Sub\n';
		expect(codes(noRedim).some((c) => c.startsWith('unallocated-dynamic-array-access'))).toBe(true);
		expect(codes(ifAccess).some((c) => c.startsWith('unallocated-dynamic-array-access'))).toBe(true);
	});
});

describe('ReDim ... As Type clause', () => {
	it.each(['Collection', 'Object', 'String'])('does not flag the type name in ReDim x(...) As %s', (typeName) => {
		const body = `Sub T()\n    Dim rowsByIdx() As ${typeName}\n    Dim cap As Long\n    cap = 16\n    ReDim rowsByIdx(1 To cap) As ${typeName}\nEnd Sub\n`;
		expect(codes(body)).not.toContain(`undeclared-variable:${typeName}`);
	});

	it('does not flag ReDim Preserve ... As Collection', () => {
		const body = 'Sub T()\n    Dim rowObjs() As Collection\n    ReDim rowObjs(1 To 8) As Collection\n    ReDim Preserve rowObjs(1 To 16) As Collection\nEnd Sub\n';
		expect(codes(body)).not.toContain('undeclared-variable:Collection');
	});

	it('still flags an undeclared variable in the ReDim bounds', () => {
		const body = 'Sub T()\n    Dim a() As Long\n    ReDim a(1 To undeclaredVar) As Long\nEnd Sub\n';
		expect(codes(body)).toContain('undeclared-variable:undeclaredVar');
	});

	it('still flags a bare use of Collection as a variable', () => {
		const body = 'Sub T()\n    Dim x As Variant\n    x = Collection\nEnd Sub\n';
		expect(codes(body)).toContain('undeclared-variable:Collection');
	});
});

describe('parenless call with a parenthesized first argument', () => {
	const ASSERT = 'Private Sub AssertTrue(ByVal condition As Boolean, ByVal message As String)\nEnd Sub\n';

	function arityCodes(call: string): string[] {
		const body = `${ASSERT}Sub T()\n    Dim v As Long\n    v = 1\n    ${call}\nEnd Sub\n`;
		return codes(body).filter((c) => c.startsWith('argument-count'));
	}

	it.each([
		'AssertTrue (v <> 0), "message"',
		'AssertTrue(v <> 0), "message"',
		'AssertTrue (v), (v = 1)',
		'AssertTrue (v) > 0, "message"',
		'AssertTrue v <> 0, "message"',
	])('counts both arguments: %s', (call) => {
		expect(arityCodes(call)).toEqual([]);
	});

	it('still flags a genuinely 1-argument call', () => {
		expect(arityCodes('AssertTrue (v <> 0)').length).toBeGreaterThan(0);
	});

	it('flags a 3-argument call with the correct count', () => {
		const body = `${ASSERT}Sub T()\n    Dim v As Long\n    v = 1\n    AssertTrue (v <> 0), "m", 1\nEnd Sub\n`;
		const src = `Option Explicit\n${body}`;
		const hit = analyzeVbaModuleSource({ source: src, moduleName: 'Module1' })
			.diagnostics.find((d) => d.code === 'argument-count');
		expect(hit?.message).toContain('got 3');
	});
});
