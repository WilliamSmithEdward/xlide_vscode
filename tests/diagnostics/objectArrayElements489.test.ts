// Diagnostics tests: elements of an array of objects that the code never set
// (issue #489). Each sample was measured through pyVBAharness on 2026-10-02
// in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode } from '../helpers/diagnostics';

const NOT_SET = 'object-variable-not-set';

function wrap(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

describe('elements of an array of objects (issue #489)', () => {
	it('reports an element never set, or set to Nothing', () => {
		const bodies = [
			['Dim a(1) As Collection', 'Main = a(0).Count'],
			['Dim a(1) As Collection', 'a(0).Add 1'],
			['Dim a(1) As Collection', 'Set a(0) = New Collection', 'Main = a(1).Count'],
			['Dim a(1) As Collection', 'Set a(0) = New Collection', 'Set a(0) = Nothing', 'Main = a(0).Count'],
			['Dim a(1) As Collection', 'Dim x As Variant', 'For Each x In a', '    Main = x.Count', 'Next'],
			['Dim a(1) As Object', 'Main = a(0).Count'],
			['Dim a() As Collection', 'ReDim a(2)', 'Main = a(1).Count'],
			['Dim a() As Collection', 'ReDim a(2)', 'Set a(1) = New Collection', 'ReDim a(2)', 'Main = a(1).Count'],
		];
		for (const body of bodies) {
			expect(byCode(analyzeModule(wrap(...body)), NOT_SET), body.join(' / ')).toHaveLength(1);
		}
	});

	it('stays quiet once the element is set, or may be', () => {
		const bodies = [
			['Dim a(1) As Collection', 'Set a(0) = New Collection', 'Main = a(0).Count'],
			['Dim a(1) As Collection', 'Dim i As Long', 'For i = 0 To 1', '    Set a(i) = New Collection', 'Next', 'Main = a(1).Count'],
			['Dim a(1) As Collection', 'Main = a(0) Is Nothing'],
			['Dim a() As Collection', 'ReDim a(2)', 'Set a(1) = New Collection', 'ReDim Preserve a(3)', 'Main = a(1).Count'],
			['Dim a(1) As Collection', 'Dim x As Variant', 'Set a(0) = New Collection', 'Set a(1) = New Collection', 'For Each x In a', '    Main = x.Count', 'Next'],
			['Dim a(1) As Collection', 'If Rnd() < 2 Then Set a(0) = New Collection', 'Main = a(0).Count'],
			['Dim a(1) As New Collection', 'Main = a(0).Count'],
		];
		for (const body of bodies) {
			expect(byCode(analyzeModule(wrap(...body)), NOT_SET), body.join(' / ')).toHaveLength(0);
		}
	});
});
