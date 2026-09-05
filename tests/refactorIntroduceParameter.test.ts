// Introduce Parameter (issue #69). The refusal that carries the weight: the
// initialiser has to mean the same thing at a call site as it did inside the
// procedure. A name that is local, or Private to the module, does not - it
// would fail to compile there, or bind to a different name that happens to
// exist. Those are refused by name.

import { describe, expect, it } from 'vitest';
import { introduceParameter } from '../src/analyzer/refactor/introduceParameter';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';

function run(source: string, name: string, others?: Record<string, string>) {
	const offset = source.indexOf(name);
	expect(offset, `'${name}' is not in the fixture`).toBeGreaterThan(-1);
	return introduceParameter({
		source,
		offset,
		moduleName: 'Module1',
		...(others ? { otherModuleSources: others } : {}),
	});
}

function applied(source: string, name: string, others?: Record<string, string>): string {
	const result = run(source, name, others);
	if (!result.ok) { throw new Error(`refused: ${result.reason}`); }
	return applyVbaTextEdits(source, result.edits);
}

function reason(source: string, name: string): string {
	const result = run(source, name);
	if (result.ok) { throw new Error(`expected a refusal, got ${result.title}`); }
	return result.reason;
}

const SIMPLE = [
	'Public Sub Report()',
	'    Dim limit As Long',
	'    limit = 3',
	'    Debug.Print limit',
	'End Sub',
	'',
].join('\r\n');

describe('what it writes', () => {
	it('adds the parameter and drops the declaration and assignment', () => {
		expect(applied(SIMPLE, 'limit')).toBe([
			'Public Sub Report(ByVal limit As Long)',
			'    Debug.Print limit',
			'End Sub',
			'',
		].join('\r\n'));
	});

	it('appends after the parameters a procedure already has', () => {
		const source = [
			'Public Sub Report(ByVal title As String)',
			'    Dim limit As Long',
			'    limit = 3',
			'    Debug.Print title, limit',
			'End Sub',
			'',
		].join('\r\n');
		expect(applied(source, 'limit'))
			.toContain('Public Sub Report(ByVal title As String, ByVal limit As Long)');
	});

	it('gives brackets to a procedure declared without them', () => {
		const source = [
			'Public Sub Report',
			'    Dim limit As Long',
			'    limit = 3',
			'    Debug.Print limit',
			'End Sub',
			'',
		].join('\r\n');
		expect(applied(source, 'limit')).toContain('Public Sub Report(ByVal limit As Long)');
	});

	it('passes the value at a bare call site in the same module', () => {
		const source = SIMPLE + [
			'Public Sub Caller()',
			'    Report',
			'End Sub',
			'',
		].join('\r\n');
		expect(applied(source, 'limit')).toContain('    Report 3\r\n');
	});

	it('passes it inside the brackets a call site already uses', () => {
		const source = SIMPLE + [
			'Public Sub Caller()',
			'    Call Report',
			'    Call Report()',
			'End Sub',
			'',
		].join('\r\n');
		const out = applied(source, 'limit');
		expect(out).toContain('    Call Report 3\r\n');
		expect(out).toContain('    Call Report(3)\r\n');
	});

	it('reports call sites in other modules separately', () => {
		const others = {
			Module2: 'Public Sub Elsewhere()\r\n    Report\r\nEnd Sub\r\n',
			Module3: 'Public Sub Nothing()\r\nEnd Sub\r\n',
		};
		const result = run(SIMPLE, 'limit', others);
		if (!result.ok) { throw new Error(result.reason); }
		expect(result.otherModules?.map((m) => m.moduleName)).toEqual(['Module2']);
		expect(applyVbaTextEdits(others.Module2, result.otherModules![0].edits))
			.toContain('    Report 3\r\n');
	});

	it('takes the declared type onto the parameter', () => {
		const source = [
			'Public Sub Report()',
			'    Dim title As String',
			'    title = "x"',
			'    Debug.Print title',
			'End Sub',
			'',
		].join('\r\n');
		expect(applied(source, 'title')).toContain('(ByVal title As String)');
	});
});

describe('what it refuses, and why', () => {
	function local(body: string[]): string {
		return ['Public Sub Report()', ...body, 'End Sub', ''].join('\r\n');
	}

	it('refuses a value that names another local', () => {
		const source = local([
			'    Dim base As Long',
			'    base = 1',
			'    Dim limit As Long',
			'    limit = base + 1',
			'    Debug.Print limit',
		]);
		expect(reason(source, 'limit')).toMatch(/names 'base'.*cannot see/s);
	});

	it('refuses a value that names a parameter', () => {
		const source = [
			'Public Sub Report(ByVal n As Long)',
			'    Dim limit As Long',
			'    limit = n',
			'    Debug.Print limit',
			'End Sub',
			'',
		].join('\r\n');
		expect(reason(source, 'limit')).toMatch(/names 'n'/);
	});

	it('refuses a value that names something Private to the module', () => {
		const source = 'Private Secret As Long\r\n\r\n' + local([
			'    Dim limit As Long',
			'    limit = Secret',
			'    Debug.Print limit',
		]);
		expect(reason(source, 'limit')).toMatch(/names 'Secret'/);
	});

	it('allows a value that names a Public module member', () => {
		const source = 'Public Shared As Long\r\n\r\n' + local([
			'    Dim limit As Long',
			'    limit = Shared',
			'    Debug.Print limit',
		]);
		expect(applied(source, 'limit')).toContain('(ByVal limit As Long)');
	});

	it('refuses an object assigned with Set, which cannot be passed ByVal', () => {
		const source = local([
			'    Dim sheet As Worksheet',
			'    Set sheet = ActiveSheet',
			'    Debug.Print sheet.Name',
		]);
		expect(reason(source, 'sheet')).toMatch(/Set, which cannot be passed ByVal/);
	});

	it('refuses a local assigned more than once', () => {
		const source = local([
			'    Dim limit As Long',
			'    limit = 1',
			'    limit = 2',
			'    Debug.Print limit',
		]);
		expect(reason(source, 'limit')).toMatch(/assigned 2 times/);
	});

	it('refuses a name that is already a parameter', () => {
		const source = [
			'Public Sub Report(ByVal limit As Long)',
			'    Debug.Print limit',
			'End Sub',
			'',
		].join('\r\n');
		expect(reason(source, 'limit')).toMatch(/already a parameter/);
	});

	it('refuses Static and Const, which a caller cannot supply', () => {
		expect(reason(local(['    Static limit As Long', '    limit = 1', '    Debug.Print limit']), 'limit'))
			.toMatch(/Static/);
		expect(reason(local(['    Const limit As Long = 1', '    Debug.Print limit']), 'limit'))
			.toMatch(/Const/);
	});
});
