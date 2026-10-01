// Diagnostics tests: Split and Filter with no compare argument take the
// module's Option Compare (issues #353 and #405). Measured in Excel 16.0
// (build 20326, 2026-10-01): under Option Compare Text, Split("aXbxc", "x")
// has three parts and Filter(Array("Ab", "cd"), "a") keeps "Ab"; under
// Binary, two parts and none. Access modules say Option Compare Database,
// which ignores case the same way.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

const CODE = 'array-subscript-out-of-bounds';

function source(option: string, ...lines: string[]): string {
	return `Option Explicit\n${option}Function Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

const SPLIT = ['Dim p', 'p = Split("aXbxc", "x")', 'Main = p(2)'];
const FILTER = ['Dim f', 'f = Filter(Array("Ab", "cd"), "a")', 'Main = f(0)'];
const INLINE = ['Main = Split("aXbxc", "x")(2)'];
// Read past a label, where the procedure-wide shape decides.
const LABELLED = ['Dim p', 'p = Split("aXbxc", "x")', 'GoTo L', 'L:', 'Main = p(2)'];

describe('Split and Filter with no compare argument', () => {
	it('ignore case under Option Compare Text or Database', () => {
		for (const option of ['Option Compare Text\n', 'Option Compare Database\n']) {
			for (const lines of [SPLIT, FILTER, INLINE, LABELLED]) {
				expect(byCode(analyzeModule(source(option, ...lines)), CODE), `${option}${lines.join(' : ')}`).toEqual([]);
			}
		}
	});

	it('match case under Option Compare Binary or none', () => {
		for (const option of ['Option Compare Binary\n', '']) {
			for (const [lines, span] of [[SPLIT, '2'], [FILTER, '0'], [INLINE, '2'], [LABELLED, '2']] as const) {
				const src = source(option, ...lines);
				expectDiagnostic(src, byCode(analyzeModule(src), CODE), CODE, { span });
			}
		}
	});

	it('follow an explicit compare argument over the module', () => {
		const binary = source('Option Compare Text\n', 'Dim p', 'p = Split("aXbxc", "x", -1, vbBinaryCompare)', 'Main = p(2)');
		expectDiagnostic(binary, byCode(analyzeModule(binary), CODE), CODE, { span: '2' });
		const filter = source('Option Compare Text\n', 'Dim f', 'f = Filter(Array("Ab", "cd"), "a", True, vbBinaryCompare)', 'Main = f(0)');
		expectDiagnostic(filter, byCode(analyzeModule(filter), CODE), CODE, { span: '0' });
		const text = source('', 'Dim p', 'p = Split("aXbxc", "x", -1, vbTextCompare)', 'Main = p(2)');
		expect(byCode(analyzeModule(text), CODE)).toEqual([]);
	});
});
