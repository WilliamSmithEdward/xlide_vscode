// Diagnostics tests: an array of a project class with no default member,
// ReDim'd and indexed, reads no default member (issue #738). Measured on
// 2026-10-03 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';
import { buildVbaProjectIndex, projectAnalysisOptionsForModule, projectProcedureSignatures } from '../../src/vbaProjectAnalysis';

function errors(body: string): string[] {
	const main = `Option Explicit\nFunction Main() As Variant\n    ${body}\nEnd Function\n`;
	const project = buildVbaProjectIndex([
		{ moduleName: 'Module1', type: 'standard', source: main },
		{ moduleName: 'K1', type: 'class', source: 'Option Explicit\nPublic Name As String\n' },
	]);
	return analyzeVbaModuleSource({ source: main, moduleName: 'Module1', moduleKind: 'standard', ...projectAnalysisOptionsForModule(project, 'Module1', projectProcedureSignatures(project)) } as never)
		.diagnostics.filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('an array of a project class (issue #738)', () => {
	it('stays quiet on ReDim, ReDim Preserve and an element', () => {
		expect(errors('Dim a() As K1\n    ReDim a(1 To 16)\n    Main = UBound(a)')).toEqual([]);
		expect(errors('Dim a() As K1\n    ReDim a(1 To 2)\n    ReDim Preserve a(1 To 2 * UBound(a))\n    Main = UBound(a)')).toEqual([]);
		expect(errors('Dim a(2) As K1\n    Main = (a(1) Is Nothing)')).toEqual([]);
	});

	it('still reports an index on an instance with no default member', () => {
		expect(errors('Dim k As K1\n    Set k = New K1\n    Main = k(1)')).toEqual(['object-default-value']);
	});
});
