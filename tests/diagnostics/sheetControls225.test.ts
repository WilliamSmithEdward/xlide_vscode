// Diagnostics tests: a worksheet's module whose class and ActiveX controls the
// workbook supplied (issue #225). Me.Nope inside Sheet1 and Sheet1.Nope()
// from a standard module are "Method or data member not found", measured in
// 64-bit Excel 16.0 through pyVBAharness.

import { describe, it, expect } from 'vitest';
import { byCode } from '../helpers/diagnostics';
import { analyzeProjectModule, type ProjectTestModule } from './helpers';

const SHEET = 'Option Explicit\nPublic Function Tag() As Long\nEnd Function\n';

function findings(currentModule: string, body: string, known = true): string[] {
	const sheetFacts = known ? { designerClass: 'Excel.Worksheet', implicitMembers: [{ name: 'CommandButton1', type: 'Object' }] } : {};
	const modules: ProjectTestModule[] = [
		{ moduleName: 'Sheet1', source: SHEET, moduleKind: 'document', ...sheetFacts },
		{ moduleName: 'Module1', source: 'Option Explicit\n' },
	];
	const src = currentModule === 'Sheet1'
		? `${SHEET}Public Sub Probe()\n${body}\nEnd Sub\n`
		: `Option Explicit\nPublic Sub Probe()\n${body}\nEnd Sub\n`;
	const extra = currentModule === 'Sheet1' ? { host: 'excel', moduleKind: 'document', ...sheetFacts } : { host: 'excel' };
	const diagnostics = analyzeProjectModule(src, modules, currentModule, extra as never);
	return byCode(diagnostics, 'member-not-found').map((d) => d.message);
}

describe("a worksheet's members, with its controls known (issue #225)", () => {
	it('reports a name the sheet does not have', () => {
		expect(findings('Sheet1', '    Debug.Print Me.Nope')).toEqual([expect.stringContaining('Nope')]);
		expect(findings('Module1', '    Debug.Print Sheet1.Nope()')).toEqual([expect.stringContaining('Nope')]);
	});

	it('finds the sheet, its code and its controls', () => {
		const body = ['    Debug.Print Me.Range("A1").Value', '    Debug.Print Me.Name', '    Debug.Print Me.Tag()', '    Debug.Print Me.CommandButton1.Caption'].join('\n');
		expect(findings('Sheet1', body)).toEqual([]);
		const outside = ['    Debug.Print Sheet1.Range("A1").Value', '    Debug.Print Sheet1.Tag()', '    Debug.Print Sheet1.CommandButton1.Caption', '    Debug.Print Sheet1.CodeName'].join('\n');
		expect(findings('Module1', outside)).toEqual([]);
	});

	it('stays quiet when the workbook supplied nothing', () => {
		expect(findings('Sheet1', '    Debug.Print Me.Nope', false)).toEqual([]);
		expect(findings('Module1', '    Debug.Print Sheet1.Nope()', false)).toEqual([]);
	});
});
