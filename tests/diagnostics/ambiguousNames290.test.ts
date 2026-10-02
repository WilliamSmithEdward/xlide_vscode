// Diagnostics tests: the same Public name in two modules, used bare (issue
// #290). Measured in Excel 16.0 64-bit (2026-10-02) through pyVBAharness:
// each is "Ambiguous name detected" while compiling.

import { describe, it, expect } from 'vitest';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';
import { buildVbaProjectIndex, projectAnalysisOptionsForModule, projectProcedureSignatures } from '../../src/vbaProjectAnalysis';

function found(body: string, module2: string, module3: string): string[] {
	const modules = [
		{ moduleName: 'Module1', type: 'standard', source: `Option Explicit\nFunction Main() As Variant\n    ${body}\nEnd Function\n` },
		{ moduleName: 'Module2', type: 'standard', source: `Option Explicit\n${module2}\n` },
		{ moduleName: 'Module3', type: 'standard', source: `Option Explicit\n${module3}\n` },
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

const FN2 = 'Public Function Foo() As Long\n    Foo = 2\nEnd Function';
const FN3 = 'Public Function Foo() As Long\n    Foo = 3\nEnd Function';

describe('a Public name two modules declare', () => {
	it('is ambiguous read bare, whatever each declares', () => {
		expect(found('Main = Foo()', FN2, FN3)).toEqual(['ambiguous-project-procedure']);
		expect(found('Main = gX', 'Public gX As Long', 'Public gX As Long')).toEqual(['ambiguous-project-procedure']);
		expect(found('Main = KK', 'Public Const KK As Long = 1', 'Public Const KK As Long = 2')).toEqual(['ambiguous-project-procedure']);
		expect(found('Main = Foo', 'Public Foo As Long', FN3)).toEqual(['ambiguous-project-procedure']);
	});

	it('compiles qualified, shadowed, or declared Private in one', () => {
		expect(found('Main = Module2.Foo()', FN2, FN3)).toEqual([]);
		expect(found('Dim Foo As Long\n    Main = Foo', FN2, FN3)).toEqual([]);
		expect(found('Main = Foo()', FN2, 'Private Function Foo() As Long\n    Foo = 3\nEnd Function')).toEqual([]);
	});
});
