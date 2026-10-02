// Diagnostics tests: late-bound calls to a known class or Collection with
// arguments the target refuses (issue #485). Each sample was measured
// through pyVBAharness on 2026-10-02 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeProjectModule } from './helpers';

const CLASS1 = 'Option Explicit\nPublic Function M(ByVal a As Long, Optional ByVal b As Long) As Long\n    M = a + b\nEnd Function\n';

function found(...lines: string[]): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
	return analyzeProjectModule(src, [
		{ moduleName: 'Module1', source: src },
		{ moduleName: 'Class1', source: CLASS1, type: 'class' },
	], 'Module1').filter((d) => d.severity === 'error').map((d) => /Run-time error '(\d+)'/.exec(d.message)?.[1] ?? d.code ?? '');
}

const K = ['Dim o As Object', 'Set o = New Class1'];
const C = ['Dim o As Object', 'Set o = New Collection'];

describe('late-bound arguments a known target refuses (issue #485)', () => {
	it('reports a named argument it lacks, too many, too few, and an argument to Count', () => {
		const cases: Array<[string[], string]> = [
			[[...K, 'Main = o.M(a:=1, c:=2)'], '448'],
			[[...K, 'Main = o.M(1, 2, 3)'], '450'],
			[[...K, 'Main = o.M()'], '449'],
			[[...C, 'o.Add Itemz:=1'], '448'],
			[[...C, 'o.Add 1, Key:="a", Befor:=1'], '448'],
			[[...C, 'o.Add 1, "a", , , 5'], '450'],
			[[...C, 'Main = o.Count(1)'], '451'],
			[['Dim v As Variant', 'Set v = New Class1', 'Main = v.M(a:=1, c:=2)'], '448'],
		];
		for (const [lines, error] of cases) {
			expect(found(...lines), lines.join(' / ')).toEqual([error]);
		}
	});

	it('stays quiet on arguments the target takes', () => {
		for (const lines of [[...K, 'Main = o.M(a:=1, b:=2)'], [...K, 'Main = o.M(1)'], [...C, 'o.Add Item:=1, Key:="a"', 'Main = o.Count'], [...C, 'o.Add 1, "a"', 'Main = o.Item("a")']]) {
			expect(found(...lines), lines.join(' / ')).toEqual([]);
		}
	});
});
