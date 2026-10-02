// Diagnostics tests: a Collection's item assigned, and members no open type
// has (issue #305). Each sample was run or compiled through pyVBAharness on
// 2026-10-02 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { getWordObjectModel } from '../../src/analyzer/host/wordObjectModel';
import { byCode } from '../helpers/diagnostics';
import { analyzeProjectModule } from './helpers';

function module(body: string): string {
	return `Option Explicit\nFunction Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
}

function codes(body: string): string[] {
	return analyzeModule(module(body)).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a Collection item assigned (issue #305)', () => {
	it('raises 424 on a number or a string item', () => {
		expect(codes('Dim c As New Collection\n    c.Add 1\n    c(1) = 5')).toEqual(['variant-value-misuse']);
		expect(codes('Dim c As New Collection\n    c.Add 1\n    c.Item(1) = 5')).toEqual(['variant-value-misuse']);
		expect(codes('Dim c As Collection\n    Set c = New Collection\n    c.Add "a"\n    c(1) = 5')).toEqual(['variant-value-misuse']);
	});

	it('leaves an object item, which takes the value through its default member', () => {
		expect(codes('Dim c As New Collection\n    c.Add Range("A1")\n    c(1) = 5')).toEqual([]);
	});

	it('refuses Count as an assignment target', () => {
		expect(codes('Dim c As New Collection\n    c.Add 1\n    c.Count = 2')).toEqual(['readonly-member-assignment']);
	});
});

describe('a member an open type does not have (issue #305)', () => {
	it('raises 438', () => {
		const bodies = [
			'Dim c As New Collection\n    c.Add 1\n    Main = c.Nope',
			'Dim c As New Collection\n    c.Nope 1',
			'Main = Cells.Nope',
			'Main = Range("A1").Nope',
			'Main = ActiveSheet.Nope',
			'Dim a As Application\n    Set a = Application\n    Main = a.Nope',
			'Dim r As Range\n    Set r = Range("A1")\n    Main = r.Nope',
		];
		for (const body of bodies) {
			expect(codes(body), body).toEqual(['runtime-member-not-found']);
		}
	});

	it('stays quiet on members the type has', () => {
		const bodies = [
			'Dim c As New Collection\n    c.Add 1\n    Main = c.Count + c.Item(1)',
			'Main = Range("A1").Address & Cells(1, 1).Value',
			'Main = ActiveSheet.Name & ActiveSheet.ChartArea.Width',
			'Dim a As Application\n    Set a = Application\n    Main = a.Version & a.Sum(1, 2)',
		];
		for (const body of bodies) {
			expect(codes(body), body).toEqual([]);
		}
	});

	it('takes a member a document module declares as one ActiveSheet may have', () => {
		const caller = module('ActiveSheet.Refresh2');
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Sheet1', moduleKind: 'document', source: 'Option Explicit\nPublic Sub Refresh2()\nEnd Sub\n' },
		], 'Module1');
		expect(byCode(diagnostics, 'runtime-member-not-found')).toHaveLength(0);
	});
});

describe('a name after Excel. (issue #305)', () => {
	it('is Method or data member not found when the library has none', () => {
		for (const name of ['Nope', 'Version', 'ScreenUpdating', 'msoTrue']) {
			expect(codes(`Main = Excel.${name}`), name).toEqual(['member-not-found']);
		}
	});

	it('takes a type, a constant or a Global member', () => {
		expect(codes('Main = Excel.xlUp + Excel.rgbHotPink\n    Main = Excel.ActiveSheet.Name & Excel.Range("A1").Address\n    Dim r As Excel.Range')).toEqual([]);
	});

	it('leaves a project that declares Excel, and the other hosts', () => {
		expect(byCode(analyzeModule(module('Dim Excel As Object\n    Set Excel = Application\n    Main = Excel.Nope')), 'member-not-found')).toEqual([]);
		expect(analyzeModule(module('Main = Excel.Nope'), { hostModel: getWordObjectModel() }).filter((diag) => diag.code === 'member-not-found')).toEqual([]);
	});
});
