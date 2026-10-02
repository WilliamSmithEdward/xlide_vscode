// Diagnostics tests: late-bound objects created from a ProgID (issue #477).
// Each sample was measured through pyVBAharness on 2026-10-02 in Excel 16.0
// (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(...lines: string[]): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
	return analyzeModule(src).filter((d) => d.severity === 'error').map((d) => /Run-time error '(-?\d+)'/.exec(d.message)?.[1] ?? d.code ?? '');
}

const RE = ['Dim re As Object', 'Set re = CreateObject("VBScript.RegExp")'];
const FSO = ['Dim fso As Object, ts As Object, p As String', 'Set fso = CreateObject("Scripting.FileSystemObject")', 'p = "x.txt"'];

describe('objects from a ProgID (issue #477)', () => {
	it('reports a ProgID no class has', () => {
		for (const progId of ['CreateObject("Scripting.Dictionery")', 'CreateObject("Dictionary")', 'CreateObject("")', 'GetObject(, "Excel.Aplication")']) {
			expect(errors('Dim o As Object', `Set o = ${progId}`), progId).toEqual(['429']);
		}
		expect(errors('Dim o As Object', 'Set o = CreateObject("scripting.dictionary")', 'Main = o.Count')).toEqual([]);
		expect(errors('Dim o As Object', 'Set o = CreateObject("MSXML2.DOMDocument.6.0")')).toEqual([]);
	});

	it('reports what a RegExp refuses', () => {
		const cases: Array<[string[], string]> = [
			[['re.Pattern = "(a"', 'Main = re.Test("a")'], '5020'],
			[['re.Pattern = "[a"', 'Main = re.Test("a")'], '5019'],
			[['re.Pattern = "*a"', 'Main = re.Test("a")'], '5018'],
			[['re.Pattern = "(?<=a)b"', 'Main = re.Test("ab")'], '5017'],
			[['Main = re.Test(Null)'], '13'],
			[['re.Global = "abc"'], '13'],
			[['re.Patern = "a"'], '438'],
		];
		for (const [lines, error] of cases) {
			expect(errors(...RE, ...lines), lines.join(' / ')).toEqual([error]);
		}
		for (const lines of [['re.Pattern = "(?:a)+b*"', 'Main = re.Test("aab")'], ['re.Pattern = "z"', 'Main = re.Execute("abc").Count'], ['re.Global = True', 're.Pattern = "[)(]"', 'Main = re.Test("(")']]) {
			expect(errors(...RE, ...lines), lines.join(' / ')).toEqual([]);
		}
	});

	it('reports an IOMode OpenTextFile does not take', () => {
		expect(errors(...FSO, 'Set ts = fso.OpenTextFile(p, 3, True)')).toEqual(['5']);
		expect(errors(...FSO, 'Set ts = fso.OpenTextFile(p, 8, True)', 'ts.Close')).toEqual([]);
	});
});
