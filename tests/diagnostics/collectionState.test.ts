// Diagnostics tests: a VBA.Collection whose contents the code makes plain
// (issue #121). Each raising sample was measured in Excel 16.0 (build 20326,
// 2026-09-26); each quiet one runs there.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

const INDEX = 'collection-index-out-of-range';
const KEY = 'collection-key-not-found';
const DUP = 'collection-key-in-use';

function wrap(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

describe('collection-index-out-of-range (issue #121)', () => {
	it('reports any index into a collection with nothing added as error 5', () => {
		for (const read of ['Main = c(1)', 'Main = c(0)', 'Main = c(-1)', 'c.Remove 1', 'Main = c.Item(1)']) {
			const src = wrap('Dim c As New Collection', read, 'Main = 1');
			expectDiagnostic(src, analyzeModule(src), INDEX, { message: ['holds nothing', "'5'"] });
		}
	});

	it('reports an index outside 1..Count after Adds as error 9', () => {
		const cases: Array<[string[], string]> = [
			[['c.Add 1', 'Main = c(0)'], '0'],
			[['c.Add 1', 'Main = c(-1)'], '-1'],
			[['c.Add 1', 'Main = c.Item(2)'], '2'],
			[['c.Add 1', 'c.Add 2', 'Main = c(3)'], '3'],
			[['c.Add 1', 'c.Remove 0'], '0'],
			[['c.Add 1', 'c.Remove 2'], '2'],
		];
		for (const [lines, span] of cases) {
			const src = wrap('Dim c As New Collection', ...lines, 'Main = 1');
			expectDiagnostic(src, analyzeModule(src), INDEX, { span, message: ['indexed 1 to', "'9'"] });
		}
	});

	it('follows Set c = New Collection, Remove, and stops at anything it cannot follow', () => {
		const set = wrap('Dim c As Collection', 'Set c = New Collection', 'Main = c(1)');
		expectDiagnostic(set, analyzeModule(set), INDEX, { message: "'5'" });
		const removed = wrap('Dim c As New Collection', 'c.Add 1', 'c.Add 2', 'c.Remove 1', 'Main = c(2)');
		expectDiagnostic(removed, analyzeModule(removed), INDEX, { span: '2', message: 'holds 1 element' });
		const quiet = wrap(
			'Dim c As New Collection, d As New Collection, e As New Collection, i As Long',
			'c.Add 1',
			'c.Add 2',
			'Main = c(2)',
			'If Main Then c.Add 3',
			'Main = c(3)',
			'Fill d',
			'Main = d(1)',
			'For i = 1 To 3',
			'    e.Add i',
			'Next',
			'Main = e(3)',
		) + 'Sub Fill(ByRef x As Collection)\n    x.Add 1\nEnd Sub\n';
		expect(byCode(analyzeModule(quiet), INDEX)).toHaveLength(0);
	});
});

describe('collection-key-not-found and collection-key-in-use (issue #121)', () => {
	it('reports a key that was never added, or was removed', () => {
		const cases: string[][] = [
			['c.Add 1, "k"', 'Main = c("nokey")'],
			['c.Add 1, "k"', 'Main = c.Item("nokey")'],
			['c.Add 1, "k"', 'c.Remove "x"'],
			['c.Add 1, "k"', 'c.Remove 1', 'Main = c("k")'],
			['c.Remove "x"'],
		];
		for (const lines of cases) {
			const src = wrap('Dim c As New Collection', ...lines, 'Main = 1');
			expectDiagnostic(src, analyzeModule(src), KEY, { message: "'5'" });
		}
	});

	it('reports a key added twice, comparing without case', () => {
		const src = wrap('Dim c As New Collection', 'c.Add 1, "k"', 'c.Add 2, "K"', 'Main = 1');
		expectDiagnostic(src, analyzeModule(src), DUP, { span: '"K"', message: "'457'" });
	});

	it('stays quiet when the key is present, or when a key it could not read was added', () => {
		const src = wrap(
			'Dim c As New Collection, d As New Collection, k As String',
			'c.Add 1, "k"',
			'Main = c("K")',
			'd.Add 1, k',
			'Main = d("anything")',
			'd.Add 2, "z"',
		);
		expect(byCode(analyzeModule(src), KEY)).toHaveLength(0);
		expect(byCode(analyzeModule(src), DUP)).toHaveLength(0);
	});
});
