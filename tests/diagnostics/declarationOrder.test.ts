// Module-level declarations that refer ahead or around a cycle (issue #211).
// Each verdict measured in 64-bit Excel 16.0.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, spanText } from '../helpers/diagnostics';
import { analyzeProjectModule } from './helpers';

const FORWARD = 'declaration-forward-reference';
const CIRCULAR = 'circular-declaration-dependency';
const MAIN = 'Public Function Main() As Variant\n    Main = 1\nEnd Function\n';

describe('declaration-forward-reference (issue #211)', () => {
	const hits = (decls: string, code = FORWARD) => {
		const src = `Option Explicit\n${decls}${MAIN}`;
		return byCode(analyzeModule(src), code).map((d) => spanText(src, d));
	};

	it.each([
		['a Const using a later Const', 'Private Const A As Long = B + 1\nPrivate Const B As Long = 1\n', 'B'],
		['an Enum member using a later member', 'Private Enum E\n    a = b + 1\n    b = 1\nEnd Enum\n', 'b'],
		['an Enum member using a later Const', 'Private Enum E\n    a = K\nEnd Enum\nPrivate Const K As Long = 1\n', 'K'],
		['a Const using a later Enum member', 'Private Const K As Long = eB\nPrivate Enum E\n    eB = 1\nEnd Enum\n', 'eB'],
		['a Const using a later Enum member, qualified', 'Private Const K As Long = E.eB\nPrivate Enum E\n    eB = 1\nEnd Enum\n', 'eB'],
		['an Enum member using a later Enum', 'Private Enum E1\n    a = eB\nEnd Enum\nPrivate Enum E2\n    eB = 1\nEnd Enum\n', 'eB'],
		['a Type member String * a later Const', 'Private Type T\n    s As String * N\nEnd Type\nPrivate Const N As Long = 5\n', 'N'],
		['an array bound using a later Const', 'Private m(N) As Long\nPrivate Const N As Long = 5\n', 'N'],
		['a String * a later Const', 'Private s As String * N\nPrivate Const N As Long = 5\n', 'N'],
		['a string Const using a later one', 'Private Const S As String = T & "x"\nPrivate Const T As String = "a"\n', 'T'],
		['a Public Const using a later Private one', 'Public Const A As Long = B\nPrivate Const B As Long = 1\n', 'B'],
		['a Const using itself', 'Private Const A As Long = A + 1\n', 'A'],
		['a Const cycle', 'Private Const A As Long = B\nPrivate Const B As Long = A\n', 'B'],
		['an Enum cycle', 'Private Enum E\n    a = b\n    b = a\nEnd Enum\n', 'b'],
		['a Type member of a later Type', 'Private Type T1\n    a As T2\nEnd Type\nPrivate Type T2\n    x As Long\nEnd Type\n', 'T2'],
		['a Type member array of a later Type', 'Private Type T1\n    a() As T2\nEnd Type\nPrivate Type T2\n    x As Long\nEnd Type\n', 'T2'],
		['a Type cycle in one module', 'Private Type T1\n    a As T2\nEnd Type\nPrivate Type T2\n    b As T1\nEnd Type\n', 'T2'],
	])('reports %s', (_name, decls, text) => {
		expect(hits(decls)).toEqual([text]);
		// Inside one module a cycle is a forward reference, not a circular dependency.
		expect(hits(decls, CIRCULAR)).toEqual([]);
	});

	it.each([
		['a Const after the one it uses', 'Private Const B As Long = 1\nPrivate Const A As Long = B + 1\n'],
		['a Type member of an earlier Type', 'Private Type T2\n    x As Long\nEnd Type\nPrivate Type T1\n    a As T2\nEnd Type\n'],
		['a variable of a later Type', 'Private v As T2\nPrivate Type T2\n    x As Long\nEnd Type\n'],
		['an Enum member using an earlier one', 'Private Enum E\n    b = 1\n    a = b + 1\nEnd Enum\n'],
	])('stays quiet for %s', (_name, decls) => {
		expect(hits(decls)).toEqual([]);
		expect(hits(decls, CIRCULAR)).toEqual([]);
	});
});

describe('circular-declaration-dependency (issue #211)', () => {
	it('reports a Type with a member of its own type', () => {
		for (const field of ['b As T', 'b() As T']) {
			const src = `Option Explicit\nPrivate Type T\n    ${field}\nEnd Type\n${MAIN}`;
			expect(byCode(analyzeModule(src), CIRCULAR).map((d) => spanText(src, d)), field).toEqual(['T']);
		}
	});

	const project = (sources: string[]) => sources.map((source, i) => ({ moduleName: `Module${i + 1}`, source }));
	const circular = (sources: string[], index: number) =>
		byCode(analyzeProjectModule(sources[index], project(sources), `Module${index + 1}`), CIRCULAR);

	it('reports a Const cycle through other modules, in each module on it', () => {
		const two = [`Public Const A1 As Long = B1 + 1\n${MAIN}`, 'Public Const B1 As Long = A1 + 1\n'];
		expect(circular(two, 0)[0].message).toContain("'A1' depends on itself through 'B1' in another module");
		expect(circular(two, 1)).toHaveLength(1);
		const three = [`Public Const A1 As Long = B1\n${MAIN}`, 'Public Const B1 As Long = C1\n', 'Public Const C1 As Long = A1\n'];
		expect([0, 1, 2].map((i) => circular(three, i).length)).toEqual([1, 1, 1]);
	});

	it('reports a Type cycle through another module', () => {
		const sources = [`Public Type T1\n    a As T2\nEnd Type\n${MAIN}`, 'Public Type T2\n    b As T1\nEnd Type\n'];
		expect(circular(sources, 0)).toHaveLength(1);
		expect(circular(sources, 1)).toHaveLength(1);
	});

	it('stays quiet across modules where there is no cycle, whatever the order', () => {
		expect(circular([`Public Const A1 As Long = B1 + 1\n${MAIN}`, 'Public Const B1 As Long = 1\n'], 0)).toHaveLength(0);
		expect(circular([`Public Const A1 As Long = Module2.B1 + 1\n${MAIN}`, 'Public Const B1 As Long = 1\n'], 0)).toHaveLength(0);
		expect(circular([`Public Type T1\n    a As T2\nEnd Type\n${MAIN}`, 'Public Type T2\n    b As Long\nEnd Type\n'], 0)).toHaveLength(0);
		// A cycle inside one module is a forward reference, in a project too.
		expect(circular([`Public Const A As Long = B\nPublic Const B As Long = A\n${MAIN}`, 'Public Const C1 As Long = 1\n'], 0)).toHaveLength(0);
	});
});
