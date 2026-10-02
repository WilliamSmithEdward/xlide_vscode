// Diagnostics tests: a project procedure named like a built-in (issue #280).
// Measured in Excel 16.0 64-bit (2026-10-02) through pyVBAharness: Module2's
// Cells, Worksheets, Range and Split take the call, and InStr never does.

import { describe, it, expect } from 'vitest';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';
import { buildVbaProjectIndex, projectAnalysisOptionsForModule, projectProcedureSignatures } from '../../src/vbaProjectAnalysis';

function found(body: string, module2: string): string[] {
	const modules = [
		{ moduleName: 'Module1', type: 'standard', source: `Option Explicit\nFunction Main() As Variant\n    ${body}\nEnd Function\n` },
		{ moduleName: 'Module2', type: 'standard', source: `Option Explicit\n${module2}\n` },
	];
	const project = buildVbaProjectIndex(modules, undefined, { conditionalCompilation: { projectConstants: {} } });
	const procedures = projectProcedureSignatures(project);
	const { diagnostics } = analyzeVbaModuleSource({
		source: modules[0].source,
		moduleName: 'Module1',
		moduleKind: 'standard',
		...projectAnalysisOptionsForModule(project, 'Module1', procedures),
	} as never);
	return diagnostics.filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a Public procedure of another module named like a built-in', () => {
	it('takes the call from Excel\'s Cells, Worksheets and Range and from VBA\'s Split', () => {
		expect(found('Main = Cells(0, 1)', 'Public Function Cells(ByVal r As Long, ByVal c As Long) As Variant\n    Cells = r + c\nEnd Function')).toEqual([]);
		expect(found('Main = Worksheets(0)', 'Public Function Worksheets(ByVal i As Long) As Variant\n    Worksheets = i\nEnd Function')).toEqual([]);
		expect(found('Main = Range("A1:")', 'Public Function Range(ByVal s As String) As Variant\n    Range = s\nEnd Function')).toEqual([]);
		expect(found('Main = Split("")(0)', 'Public Function Split(ByVal s As String) As Variant\n    Split = Array(1, 2)\nEnd Function')).toEqual([]);
	});

	it('leaves the built-in judged where no procedure takes the name', () => {
		expect(found('Main = Cells(0, 1)', 'Public Function Other() As Variant\nEnd Function')).toEqual(['host-argument-out-of-range']);
	});

	it('does not take the call from VBA\'s InStr', () => {
		expect(found('Main = InStr(0, "abc", "a")', 'Public Function InStr(ByVal a As Long, ByVal b As String, ByVal c As String) As Variant\n    InStr = 1\nEnd Function')).toEqual(['runtime-argument-value']);
	});
});
