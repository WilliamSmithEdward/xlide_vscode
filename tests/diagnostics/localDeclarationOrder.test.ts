// A procedure's own Dim, Static or Const covers only the lines after it. Each
// verdict is a full project compile in 64-bit Excel 16.0.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, spanText } from '../helpers/diagnostics';

const module = (body: string, decls = '', explicit = true): string =>
	`${explicit ? 'Option Explicit\n' : ''}${decls}Sub P1()\n${body}End Sub\n`;

describe('a local declaration used above its line', () => {
	it.each([
		['a Const value using a later Const', '    Const A As Long = B + 1\n    Const B As Long = 1\n', 'B'],
		['a Dim bound using a later Const', '    Dim a(N) As Long\n    Const N As Long = 5\n', 'N'],
		['a statement using a later Const', '    Debug.Print K\n    Const K As Long = 1\n', 'K'],
		['an assignment to a later Dim', '    x = 1\n    Dim x As Long\n', 'x'],
		['a read of a later Static', '    Debug.Print s\n    Static s As Long\n', 's'],
		['a use inside an earlier loop', '    Dim i As Long\n    For i = 1 To 2\n        Debug.Print K\n    Next\n    Const K As Long = 1\n', 'K'],
	])('is Variable not defined under Option Explicit: %s', (_name, body, name) => {
		const src = module(body);
		const hits = byCode(analyzeModule(src), 'undeclared-variable');
		expect(hits.map((d) => spanText(src, d))).toEqual([name]);
		expect(hits[0].message).toContain('declared further down the procedure');
	});

	it('is a duplicate declaration where the earlier use found a module declaration', () => {
		for (const [body, decls, name] of [
			['    Debug.Print K\n    Const K As Long = 2\n', 'Private Const K As Long = 1\n', 'K'],
			['    m = 1\n    Dim m As Long\n', 'Private m As Long\n', 'm'],
		]) {
			const src = module(body, decls);
			const hits = byCode(analyzeModule(src), 'duplicate-declaration');
			expect(hits.map((d) => spanText(src, d)), body).toEqual([name]);
			expect(byCode(analyzeModule(src), 'undeclared-variable'), body).toEqual([]);
		}
	});

	it('without Option Explicit, is a duplicate declaration, or a non-constant Const value', () => {
		const statement = module('    Debug.Print K\n    Const K As Long = 1\n', '', false);
		expect(byCode(analyzeModule(statement), 'duplicate-declaration')).toHaveLength(1);
		const chain = module('    Const A As Long = B + 1\n    Const B As Long = 1\n', '', false);
		expect(byCode(analyzeModule(chain), 'const-value-not-constant')).toHaveLength(1);
		expect(byCode(analyzeModule(chain), 'duplicate-declaration')).toHaveLength(0);
	});

	it('is fine after the line, and anywhere after a declaration inside a block', () => {
		const quiet = (body: string) => analyzeModule(module(body)).filter((d) => /undeclared-variable|duplicate-declaration/.test(d.code ?? ''));
		expect(quiet('    Const K As Long = 1\n    Debug.Print K\n')).toEqual([]);
		expect(quiet('    If True Then\n        Const K As Long = 1\n    End If\n    Debug.Print K\n')).toEqual([]);
		// A member and a named argument are not reads of the local.
		expect(quiet('    Dim o As Object\n    Debug.Print o.K\n    Const K As Long = 1\n')).toEqual([]);
		expect(quiet('    MsgBox Prompt:="x"\n    Dim Prompt As String\n')).toEqual([]);
	});
});
