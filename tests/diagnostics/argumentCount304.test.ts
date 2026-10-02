// Diagnostics tests: argument counts of a Collection's methods and of
// Excel's bare globals (issue #304). Each sample was compiled through
// pyVBAharness on 2026-10-02 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode } from '../helpers/diagnostics';
import { analyzeProjectModule } from './helpers';

const COUNT = 'argument-count';

function module(line: string): string {
	return `Option Explicit\nFunction Main() As Variant\n    Dim c As New Collection\n    c.Add 0\n    ${line}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
}

describe('argument counts the host model does not carry (issue #304)', () => {
	it('reports a call the VBE refuses', () => {
		const lines = [
			'c.Add',
			'c.Add 1, , , , 5',
			'Main = c.Item()',
			'Main = c.Count(1)',
			'Main = Cells(1, 1, 1).Address',
			'Main = Range("A1", "B2", "C3").Address',
			'Main = Evaluate()',
			'Main = Intersect(Range("A1")) Is Nothing',
			'Main = Union(Range("A1")).Address',
			'Main = Sheets.Count(1)',
		];
		for (const line of lines) {
			expect(byCode(analyzeModule(module(line)), COUNT), line).toHaveLength(1);
		}
	});

	it('stays quiet on calls that compile', () => {
		const lines = [
			'Main = Union(Range("A1"), Range("B2")).Address',
			'c.Add 1, "k"',
			'Main = c.Item(1)',
			'c.Add Item:=1, Key:="j"',
			'c.Add (1), "j"',
			'c.Remove 1',
			'Main = c.Count + c(1)',
			'Main = Range(Cells(1, 1), Cells(2, 2)).Address',
			'Main = Cells(1).Address',
			'Main = Evaluate("1+1")',
			'Main = Sheets.Count + ActiveSheet.Range("A1").Count',
		];
		for (const line of lines) {
			expect(byCode(analyzeModule(module(line)), COUNT), line).toHaveLength(0);
		}
	});

	it('leaves a project class named Collection to its own members', () => {
		const caller = 'Option Explicit\nFunction Main() As Variant\n    Dim c As New Collection\n    c.Add\n    Main = c.Count(1)\nEnd Function\n';
		const diagnostics = analyzeProjectModule(caller, [
			{
				moduleName: 'Collection',
				moduleKind: 'class',
				source: 'Option Explicit\nPublic Sub Add()\nEnd Sub\nPublic Function Count(ByVal x As Long) As Long\nEnd Function\n',
			},
		], 'Module1');
		expect(byCode(diagnostics, COUNT)).toHaveLength(0);
	});

	it('leaves a project procedure named after a global to its own signature', () => {
		const source = 'Option Explicit\nFunction Cells(a, b, c) As Long\nEnd Function\nFunction Main() As Variant\n    Main = Cells(1, 2, 3)\nEnd Function\n';
		expect(byCode(analyzeModule(source), COUNT)).toHaveLength(0);
	});
});
