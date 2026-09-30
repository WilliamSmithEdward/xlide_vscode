// Diagnostics tests: runtime errors in a block's own line (issue #233). A
// For's bounds and step, a Select Case subject, a Do, Loop or While
// condition and a With subject are evaluated like any other expression.
// Measured through pyVBAharness in Excel 16.0 on 2026-09-30: each raising
// sample raises its error, and each quiet one runs.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { parseModule } from '../../src/analyzer/parser/parseModule';
import type { BodyNode, ProcedureNode } from '../../src/analyzer/parser/nodes';
import { blockHeaderStatements } from '../../src/analyzer/diagnostics/walker';
import { byCode } from '../helpers/diagnostics';

function main(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

const ERRORS: ReadonlyArray<[string, string, string]> = [
	['CLng("abc")', 'runtime-conversion-value', '"abc"'],
	['10 / 0', 'division-by-zero', '0'],
	['Len(Left("abc", -1))', 'runtime-argument-value', '-1'],
];

const POSITIONS: ReadonlyArray<[string, string[]]> = [
	['For start', ['Dim i As Long', 'For i = {e} To 10', 'Next']],
	['For limit', ['Dim i As Long', 'For i = 1 To {e}', 'Next']],
	['For Step', ['Dim i As Long', 'For i = 1 To 10 Step {e}', 'Next']],
	['Select Case', ['Select Case {e}', 'Case Else', 'End Select']],
	['Do While', ['Do While {e} > 0', '    Exit Do', 'Loop']],
	['Loop Until', ['Do', '    Main = 1', 'Loop Until {e} > 0']],
	['While', ['While {e} > 0', 'Wend']],
	['With', ['With Range(CStr({e}))', 'End With']],
];

describe('runtime rules read block headers (issue #233)', () => {
	for (const [expression, code, marked] of ERRORS) {
		it.each(POSITIONS)(`${code} in %s`, (_label, lines) => {
			const src = main(...lines.map((line) => line.replace('{e}', expression)));
			const hits = byCode(analyzeModule(src), code);
			expect(hits).toHaveLength(1);
			expect(src.slice(hits[0].span.start, hits[0].span.end)).toContain(marked);
		});
	}

	it.each([
		['array-subscript-out-of-bounds', ['Dim a(2) As Long', 'Select Case a(5)', 'Case Else', 'End Select']],
		['array-subscript-out-of-bounds', ['Dim a(2) As Long', 'Dim i As Long', 'For i = 1 To a(5)', 'Next']],
		['object-variable-not-set', ['Dim c As Collection', 'Select Case c.Count', 'Case Else', 'End Select']],
		['object-variable-not-set', ['Dim c As Collection', 'Dim i As Long', 'For i = 1 To c.Count', 'Next']],
		['arithmetic-overflow', ['Select Case CInt(40000)', 'Case Else', 'End Select']],
		['arithmetic-overflow', ['Dim i As Long', 'For i = 1 To CInt(40000)', 'Next']],
	])('%s in a header', (code, lines) => {
		expect(byCode(analyzeModule(main(...lines)), code)).toHaveLength(1);
	});

	it.each([
		['a For whose bounds convert', ['Dim i As Long', 'For i = 1 To CLng("3")', 'Next', 'Main = i']],
		['a Select Case subject that works', ['Select Case Len("abc")', 'Case 3', '    Main = 1', 'End Select']],
		['a With whose string holds a colon', ['With Range("A1:B2")', '    Main = .Count', 'End With']],
		['a Loop condition that divides safely', ['Dim n As Long', 'Do', '    n = n + 1', 'Loop Until n / 2 > 1', 'Main = n']],
		['a one-line If opening a With', ['Dim c As New Collection', 'If True Then With c: .Add 1: End With', 'Main = c.Count']],
		['a comment after the header', ["Dim i As Long", "For i = 1 To 2 ' CLng(\"abc\") and 1 / 0", 'Next', 'Main = i']],
	])('leaves %s alone', (_label, lines) => {
		expect(analyzeModule(main(...lines)).filter((d) => d.severity === 'error')).toEqual([]);
	});

	it('reports a header once, not again inside the block', () => {
		const src = main('Dim i As Long', 'For i = 1 To CLng("abc")', '    Main = i', 'Next');
		expect(byCode(analyzeModule(src), 'runtime-conversion-value')).toHaveLength(1);
	});
});

describe('blockHeaderStatements (issue #233)', () => {
	const blocks = (src: string): BodyNode[] => {
		const proc = parseModule(src).members.find((m): m is ProcedureNode => m.kind === 'Procedure') as ProcedureNode;
		const out: BodyNode[] = [];
		const walk = (body: readonly BodyNode[]): void => {
			for (const node of body) {
				if ('body' in node && Array.isArray(node.body)) {
					out.push(node);
					walk(node.body);
				}
			}
		};
		walk(proc.body);
		return out;
	};
	const text = (src: string, node: BodyNode): { before?: string; after?: string } => {
		const { before, after } = blockHeaderStatements(src, node);
		return { before: before?.raw, after: after?.raw };
	};

	it('gives the header line and a Do\'s Loop condition', () => {
		const src = main('Do While x > 0', '    y = 1', 'Loop Until z', 'For i = 1 To 3 Step 2', 'Next', 'Select Case k', 'End Select');
		expect(blocks(src).map((node) => text(src, node))).toEqual([
			{ before: 'Do While x > 0', after: 'Loop Until z' },
			{ before: 'For i = 1 To 3 Step 2', after: undefined },
			{ before: 'Select Case k', after: undefined },
		]);
	});

	it('stops at a colon, and at a comment, but not at a colon in a string', () => {
		const src = main('If True Then With c: .Add 1: End With', "With Range(\"A1:B2\") ' note", 'End With');
		expect(blocks(src).filter((node) => node.kind === 'WithBlock').map((node) => text(src, node).before)).toEqual(['With c', 'With Range("A1:B2")']);
	});

	it('gives nothing for an If, and no Loop line when the Loop has no condition', () => {
		const src = main('If x Then', 'End If', 'Do', 'Loop');
		expect(blocks(src).map((node) => text(src, node))).toEqual([{}, { before: 'Do', after: undefined }]);
	});
});
