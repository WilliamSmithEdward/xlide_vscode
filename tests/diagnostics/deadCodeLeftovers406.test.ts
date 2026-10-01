// Diagnostics tests: code that never runs raises nothing (issue #406). A
// Const that is False, a loop of no pass, and code after `GoTo Done` that
// builds its own state. Every case was measured in Excel 16.0 (build 20326,
// 2026-10-01).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode } from '../helpers/diagnostics';

const DECL = 'Dim c As Collection, d As Long, a() As Long, i As Long, x As Variant';
const RULES = ['object-variable-not-set', 'division-by-zero', 'unallocated-dynamic-array-access', 'array-subscript-out-of-bounds', 'collection-index-out-of-range', 'collection-key-in-use'];

function source(moduleLines: string[], ...lines: string[]): string {
	return `Option Explicit\n${moduleLines.map((line) => `${line}\n`).join('')}Function Main() As Variant\n${[DECL, ...lines].map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

function found(src: string): string[] {
	const diagnostics = analyzeModule(src);
	return RULES.flatMap((code) => byCode(diagnostics, code).map(() => code));
}

describe('code that never runs', () => {
	it('is not judged under a Const that is False', () => {
		for (const [moduleLines, lines] of [
			[['Const DEBUGGING = False'], ['If DEBUGGING Then Main = c.Count', 'Main = 1']],
			[['Const DEBUGGING = False'], ['If DEBUGGING Then', '    Main = 10 / d', 'End If', 'Main = 1']],
			[['Const DEBUGGING As Boolean = False'], ['If DEBUGGING Then Main = c.Count', 'Main = 1']],
			[[], ['Const LEVEL = 0', 'If LEVEL > 0 Then Main = 10 / d', 'Main = 1']],
		] as const) {
			expect(found(source([...moduleLines], ...lines)), lines.join(' / ')).toEqual([]);
		}
	});

	it('is not judged in a loop that runs no pass', () => {
		for (const lines of [
			['Main = 1', 'For i = 1 To 0', '    Main = 10 / d', 'Next'],
			['Main = 1', 'For i = 0 To 1 Step -1', '    Main = 10 / d', 'Next'],
			['Main = 1', 'Do While False', '    Main = 10 / d', 'Loop'],
			['Main = 1', 'Do Until True', '    Main = 10 / d', 'Loop'],
			['Main = 1', 'While d <> 0', '    Main = 10 / d', 'Wend'],
			['Main = 1', 'For Each x In Array()', '    Main = 10 / d', 'Next'],
			['Main = 1', 'For i = 1 To 0', '    Main = c.Count', 'Next'],
			['Main = 1', 'For i = 1 To 0', '    Main = a(0)', 'Next'],
			['Main = 1', 'While d <> 0', '    Main = c.Count', 'Wend'],
			['Main = 1', 'For Each x In Array()', '    Main = c.Count', 'Next'],
		]) {
			expect(found(source([], ...lines)), lines.join(' / ')).toEqual([]);
		}
	});

	it('is not judged after GoTo, whatever state it builds', () => {
		for (const lines of [
			['Main = 1', 'GoTo Done', 'ReDim a(0)', 'a(2) = 1', 'Done:'],
			['Main = 1', 'GoTo Done', 'Set c = New Collection', 'c.Remove 2', 'Done:'],
			['Main = 1', 'GoTo Done', 'Set c = New Collection', 'c.Add 1, "k"', 'c.Add 2, "k"', 'Done:'],
			['Main = 1', 'GoTo Done', 'Main = c.Count', 'Done:'],
		]) {
			expect(found(source([], ...lines)), lines.join(' / ')).toEqual([]);
		}
	});

	it('is still judged where it runs', () => {
		for (const [moduleLines, lines, code] of [
			[['Const DEBUGGING = True'], ['If DEBUGGING Then Main = c.Count', 'Main = 1'], 'object-variable-not-set'],
			[[], ['Main = 1', 'For i = 1 To 1', '    Main = 10 / d', 'Next'], 'division-by-zero'],
			[[], ['Main = 1', 'Do', '    Main = 10 / d', 'Loop While False'], 'division-by-zero'],
			[[], ['Main = 1', 'For Each x In Array(1)', '    Main = 10 / d', 'Next'], 'division-by-zero'],
			[[], ['Main = 1', 'ReDim a(0)', 'a(2) = 1'], 'array-subscript-out-of-bounds'],
			[[], ['Set c = New Collection', 'c.Remove 2'], 'collection-index-out-of-range'],
		] as const) {
			expect(found(source([...moduleLines], ...lines)), lines.join(' / ')).toContain(code);
		}
	});

	it('follows a local that hides the Const', () => {
		const src = source(['Const DEBUGGING = False'], 'Dim DEBUGGING As Boolean', 'DEBUGGING = True', 'If DEBUGGING Then Main = c.Count');
		expect(found(src)).toContain('object-variable-not-set');
		const parameter = 'Option Explicit\nConst DEBUGGING = False\nFunction Main(DEBUGGING As Boolean) As Variant\n    Dim c As Collection\n    If DEBUGGING Then Main = c.Count\nEnd Function\n';
		expect(found(parameter)).toContain('object-variable-not-set');
	});
});
